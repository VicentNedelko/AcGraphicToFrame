using AcGraphicToFrame.Helpers;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AcGraphicToFrame.Managers
{
    internal static class DwgManager
    {
        private static Vector3d FrameBorderCompensation = new Vector3d(7.5, 0, 0);

        internal static string CloneDrawingToFrame(string dwgFile, ExcelWorksheet sheet)
        {
            var rootFolder = Path.GetDirectoryName(dwgFile);

            try
            {
                using (Database sourceDatabase = new Database(false, true))
                {
                    sourceDatabase.ReadDwgFile(dwgFile, FileOpenMode.OpenForReadAndAllShare, false, null);
                    sourceDatabase.CloseInput(true);

                    using (var sourceTransaction = sourceDatabase.TransactionManager.StartTransaction())
                    {
                        var sourceBlockTable = sourceTransaction.GetObject(sourceDatabase.BlockTableId, OpenMode.ForRead) as BlockTable;
                        var sourceBlockTableRecord = sourceTransaction.GetObject(sourceBlockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                        var sourceObjectIdCollection = new ObjectIdCollection();
                        var sourceExtents = new Extents3d();

                        var textHeightList = new List<int>();

                        foreach (ObjectId sourceObjectId in sourceBlockTableRecord)
                        {
                            var sourceEntity = sourceTransaction.GetObject(sourceObjectId, OpenMode.ForRead) as Entity;

                            if (sourceObjectId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(DBText))))
                            {
                                textHeightList.Add((int)Math.Truncate((double)sourceEntity.GetType().GetProperty("Height").GetValue(sourceEntity, null)));
                            }

                            sourceExtents.AddExtents(sourceEntity.GeometricExtents);
                            sourceObjectIdCollection.Add(sourceObjectId);
                        }

                        var scale = GetScaleFactor(textHeightList);

                        var sourceModelCenter = sourceExtents.MinPoint + (sourceExtents.MaxPoint - sourceExtents.MinPoint) * 0.5;
                        var sourceHeight = sourceExtents.MaxPoint.Y - sourceExtents.MinPoint.Y;
                        var sourceWidth = sourceExtents.MaxPoint.X - sourceExtents.MinPoint.X;

                        IdMapping idMapping = new IdMapping();

                        Extents3d destinationExt = new Extents3d();

                        var formatValue = FormatHelper.GetFormatValue(sourceHeight, sourceWidth, scale);

                        using (Database destinationDatabase = new Database(false, true))
                        {
                            var pathToFrame = FileManager.GetPathToFrameFile(formatValue, rootFolder, scale);
                            destinationDatabase.ReadDwgFile(pathToFrame, FileOpenMode.OpenForReadAndWriteNoShare, false, null);
                            destinationDatabase.CloseInput(true);

                            var freeRowIndex = XlsxManager.GetFreeRowIndex(sheet);
                            var indexes = GetIndexesTable(sheet, freeRowIndex);

                            using (var destinationTransaction = destinationDatabase.TransactionManager.StartTransaction())
                            {
                                var destBlockTable = destinationTransaction.GetObject(destinationDatabase.BlockTableId, OpenMode.ForRead) as BlockTable;
                                var destinationBlockTableRecord = destinationTransaction.GetObject(destBlockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;


                                foreach (ObjectId destObjectId in destinationBlockTableRecord)
                                {
                                    var destEntity = destinationTransaction.GetObject(destObjectId, OpenMode.ForRead) as Entity;

                                    if (destObjectId.ObjectClass.IsDerivedFrom(RXClass.GetClass(typeof(BlockReference))))
                                    {
                                        SetFrameIndexes(destinationTransaction, destObjectId, indexes);
                                    }

                                    destinationExt.AddExtents(destEntity.GeometricExtents);
                                }

                                var destinationModelCenter = destinationExt.MinPoint + (destinationExt.MaxPoint - destinationExt.MinPoint) * 0.5 + FrameBorderCompensation * scale;
                                // TODO: add scale factor * to compensation

                                Matrix3d matDisplacement = Matrix3d.Displacement(destinationModelCenter - sourceModelCenter);

                                foreach (ObjectId objectId in sourceBlockTableRecord)
                                {
                                    TransformDbCollection(sourceTransaction, matDisplacement, objectId);
                                }

                                sourceTransaction.Commit();

                                sourceDatabase.WblockCloneObjects(sourceObjectIdCollection, destinationBlockTableRecord.ObjectId, idMapping, DuplicateRecordCloning.Replace, false);

                                destinationTransaction.Commit();

                                var resultFileName = FileManager.GetResultFileName(indexes);
                                var filename = string.Concat(resultFileName, Constants.FileExtensionSeparator);
                                var resultFilePath = Path.Combine(rootFolder, Constants.ResultsFolderName, filename);
                                destinationDatabase.SaveAs(resultFilePath, DwgVersion.Current);

                                return "SUCCESS";
                            }
                        }
                    }
                }
            }
            catch(System.Exception ex)
            {
                return ex.Message;
            }
        }

        private static string[] GetIndexesTable(ExcelWorksheet sheet, int freeRowIndex)
        {
            var materialN = Math.Truncate((double)sheet.Cells[$"B{freeRowIndex}"].Value).ToString();
            var documentN = Math.Truncate((double)sheet.Cells[$"C{freeRowIndex}"].Value).ToString();
            var sheetN = sheet.Cells[$"D{freeRowIndex}"].Value.ToString();

            return new[] { materialN, documentN, sheetN };
        }

        private static void SetFrameIndexes(Transaction transaction, ObjectId objectId, string[] indexes)
        {
            var blockReference = transaction.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            var blockProps = blockReference.GetType().GetProperties().ToList();
            var propNameValue = blockProps.FirstOrDefault(x => x.Name == "Name")?.GetValue(blockReference).ToString();

            if (propNameValue == Constants.BlockName)
            {
                var blockAttrs = blockReference.AttributeCollection;

                foreach(ObjectId attrId in blockAttrs)
                {
                    var attrReference = transaction.GetObject(attrId, OpenMode.ForRead) as AttributeReference;
                    switch (attrReference.Tag)
                    {
                        case Constants.MATNR: UpdateAttribute(attrReference, indexes[0]); break;
                        case Constants.DOKNR: UpdateAttribute(attrReference, indexes[1]); break;
                    }
                }
            }
        }

        private static void UpdateAttribute(AttributeReference attrReference, string newValue)
        {
            attrReference.UpgradeOpen();
            attrReference.TextString = newValue;
            attrReference.DowngradeOpen();
        }

        private static void TransformDbCollection(Transaction transaction, Matrix3d matrix, ObjectId objectId)
        {
            var entity = transaction.GetObject(objectId, OpenMode.ForRead) as Entity;
            entity.UpgradeOpen();
            entity.TransformBy(matrix);
            entity.DowngradeOpen();
        }

        private static double GetScaleFactor(List<int> textHeightList)
        {
            var mostOccuringTextValue = textHeightList.GroupBy(x => x)
                .OrderByDescending(group => group.Count())
                .Select(x => x.Key)
                .FirstOrDefault();

            return Math.Truncate(mostOccuringTextValue / 2.5);
        }
    }
}
