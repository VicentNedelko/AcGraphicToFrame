using AcGraphicToFrame.Exceptions;
using AcGraphicToFrame.Helpers;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace AcGraphicToFrame.Managers
{
    internal static class DwgManager
    {
        internal static (string status, string fileName) CloneDrawingToFrame(string dwgFile, ExcelWorksheet sheet)
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

                            if (!(IsTeklaText(sourceObjectId, sourceEntity) || IsPlusLayerLine(sourceEntity) || IsWipeOutObject(sourceObjectId)))
                            {
                                sourceObjectIdCollection.Add(sourceObjectId);
                                sourceExtents.AddExtents(sourceEntity.GeometricExtents);
                            }
                        }

                        var scale = GetScaleFactor(textHeightList);

                        Point3d sourceRightUpCorner = new Point3d(sourceExtents.MaxPoint.X, sourceExtents.MaxPoint.Y, 0);
                        var sourceHeight = sourceExtents.MaxPoint.Y - sourceExtents.MinPoint.Y;
                        var sourceWidth = sourceExtents.MaxPoint.X - sourceExtents.MinPoint.X;

                        IdMapping idMapping = new IdMapping();

                        Extents3d destinationExt = new Extents3d();

                        var formatValue = FormatHelper.GetFormatValue(sourceHeight, sourceWidth, scale);

                        using (Database destinationDatabase = new Database(false, true))
                        {
                            var pathToFrame = FileManager.GetPathToFrameFile(formatValue, rootFolder, scale);
                            if (!File.Exists(pathToFrame))
                            {
                                throw new FileNotFoundException($"Frames for scale {scale} not found! Check Frames folder.");
                            }
                            destinationDatabase.ReadDwgFile(pathToFrame, FileOpenMode.OpenForReadAndWriteNoShare, false, null);
                            destinationDatabase.CloseInput(true);

                            var freeRowIndex = XlsxManager.GetFreeRowIndex(sheet);
                            var indexes = XlsxManager.GetIndexesTable(sheet, freeRowIndex);

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

                                Point3d destinationInnerRightUpCorner = new Point3d(destinationExt.MaxPoint.X - 10 * scale, destinationExt.MaxPoint.Y - 10 * scale, 0);

                                Matrix3d matDisplacement = Matrix3d.Displacement(destinationInnerRightUpCorner - sourceRightUpCorner);

                                foreach (ObjectId objectId in sourceBlockTableRecord)
                                {
                                    TransformDbCollection(sourceTransaction, matDisplacement, objectId);
                                }

                                sourceTransaction.Commit();

                                sourceDatabase.WblockCloneObjects(sourceObjectIdCollection, destinationBlockTableRecord.ObjectId, idMapping, DuplicateRecordCloning.Replace, false);

                                destinationTransaction.Commit();

                                var resultFileName = FileManager.GetResultFileName(indexes);
                                var filename = string.Concat(resultFileName, Constants.FileExtensionSeparator);
                                var resultFolderPath = Path.Combine(rootFolder, Constants.ResultsFolderName);
                                if (!Directory.Exists(resultFolderPath))
                                {
                                    Directory.CreateDirectory(resultFolderPath);
                                }
                                var resultFilePath = Path.Combine(resultFolderPath, filename);
                                destinationDatabase.SaveAs(resultFilePath, DwgVersion.Current);

                                return ("SUCCESS", dwgFile);
                            }
                        }
                    }
                }
            }
            catch (MissedCustomInfoException ex)
            {
                MessageBox.Show(ex.Message, "Protocol ERROR", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                throw;
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show(ex.Message, "Frames ERROR", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return (string.Concat("ERROR - ", ex.Message), dwgFile);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"{ex.Message}.\nCheck Protocol.xlsx for more information.", "DWG processor", MessageBoxButton.OK, MessageBoxImage.Error);
                return (string.Concat("ERROR - ", ex.Message), dwgFile);
            }
        }

        private static void SetFrameIndexes(Transaction transaction, ObjectId objectId, string[] indexes)
        {
            var blockReference = transaction.GetObject(objectId, OpenMode.ForRead) as BlockReference;
            var blockProps = blockReference.GetType().GetProperties().ToList();
            var propNameValue = blockProps.FirstOrDefault(x => x.Name == "Name")?.GetValue(blockReference).ToString();

            if (propNameValue == Constants.BlockName)
            {
                var blockAttrs = blockReference.AttributeCollection;

                foreach (ObjectId attrId in blockAttrs)
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

            return Math.Round(mostOccuringTextValue / 2.5);
        }

        private static bool IsTeklaText(ObjectId objectId, Entity entity) // add transaction
        {
            return objectId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(DBText)))
                && (string)entity.GetType().GetProperty("TextString").GetValue(entity, null) == "Tekla structures";
        }

        private static bool IsPlusLayerLine(Entity entity)
        {
            return entity.Layer == "+";
        }

        private static bool IsWipeOutObject(ObjectId objectId)
        {
            return objectId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(Wipeout)));
        }
    }
}
