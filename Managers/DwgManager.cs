using AcGraphicToFrame.Helpers;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AcGraphicToFrame.Managers
{
    internal static class DwgManager
    {
        private static Vector3d FrameBorderCompensation = new Vector3d(7.5, 0, 0);
        internal static void CloneDrawingToFrame(string file)
        {
            var rootFolder = Path.GetDirectoryName(file);
            using (Database sourceDatabase = new Database(false, true))
            {
                sourceDatabase.ReadDwgFile(file, FileOpenMode.OpenForReadAndAllShare, false, null);
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
                    var sourceLength = sourceExtents.MaxPoint.X - sourceExtents.MinPoint.X;
                    Matrix3d matScale = Matrix3d.Scaling(scale, sourceModelCenter);

                    foreach (ObjectId sourceObjectId in sourceBlockTableRecord)
                    {
                        TransformDbCollection(sourceTransaction, matScale, sourceObjectId);
                    }

                    Extents3d scaledSourceExtents = new Extents3d();
                    ObjectIdCollection scaledSourceObjIdCollection = new ObjectIdCollection();

                    foreach (ObjectId sourceObjectId in sourceBlockTableRecord)
                    {
                        var sourceEntity = sourceTransaction.GetObject(sourceObjectId, OpenMode.ForRead) as Entity;

                        scaledSourceExtents.AddExtents(sourceEntity.GeometricExtents);
                        scaledSourceObjIdCollection.Add(sourceObjectId);
                    }

                    var scaledModelCenter = scaledSourceExtents.MinPoint + (scaledSourceExtents.MaxPoint - scaledSourceExtents.MinPoint) * 0.5;
                    var scaledSourceHeight = scaledSourceExtents.MaxPoint.Y - scaledSourceExtents.MinPoint.Y;
                    var scaledSourceWidth = scaledSourceExtents.MaxPoint.X - scaledSourceExtents.MinPoint.X;

                    IdMapping idMapping = new IdMapping();

                    Extents3d destinationExt = new Extents3d();

                    var formatValue = FormatHelper.GetFormatValue(scaledSourceHeight, scaledSourceWidth);

                    using (Database destinationDatabase = new Database(false, true))
                    {
                        var pathToFrame = FileManager.GetPathToFrameFile(formatValue, rootFolder);
                        destinationDatabase.ReadDwgFile(pathToFrame, FileOpenMode.OpenForReadAndWriteNoShare, false, null);
                        destinationDatabase.CloseInput(true);

                        using (var destinationTransaction = destinationDatabase.TransactionManager.StartTransaction())
                        {
                            var destBlockTable = destinationTransaction.GetObject(destinationDatabase.BlockTableId, OpenMode.ForRead) as BlockTable;
                            var destinationBlockTableRecord = destinationTransaction.GetObject(destBlockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                            var customFrameFileName = FileManager.GetCustomFramePath(rootFolder);
                            var resultFileName = Path.GetFileNameWithoutExtension(customFrameFileName);

                            var indexes = GetCustomFrameIndexes(resultFileName);

                            foreach (ObjectId destObjectId in destinationBlockTableRecord)
                            {
                                var destEntity = destinationTransaction.GetObject(destObjectId, OpenMode.ForRead) as Entity;

                                if (destObjectId.ObjectClass.IsDerivedFrom(RXClass.GetClass(typeof(BlockReference))))
                                {
                                    SetFrameIndexes(destinationTransaction, destObjectId, indexes);
                                }

                                destinationExt.AddExtents(destEntity.GeometricExtents);
                            }

                            var destinationModelCenter = destinationExt.MinPoint + (destinationExt.MaxPoint - destinationExt.MinPoint) * 0.5 + FrameBorderCompensation;

                            Matrix3d matDisplacement = Matrix3d.Displacement(destinationModelCenter - sourceModelCenter);

                            foreach (ObjectId objectId in sourceBlockTableRecord)
                            {
                                TransformDbCollection(sourceTransaction, matDisplacement, objectId);
                            }

                            sourceTransaction.Commit();

                            sourceDatabase.WblockCloneObjects(sourceObjectIdCollection, destinationBlockTableRecord.ObjectId, idMapping, DuplicateRecordCloning.Replace, false);

                            destinationTransaction.Commit();

                            var filename = string.Concat(resultFileName, Constants.FileExtensionSeparator, Constants.DwgExtension);
                            var resultFilePath = Path.Combine(rootFolder, Constants.ResultsFolderName, filename);
                            destinationDatabase.SaveAs(resultFilePath, DwgVersion.Current);

                            FileManager.MoveCustomFrameToUsedFolder(customFrameFileName, rootFolder);
                        }
                    }
                }
            }
        }

        private static string[] GetCustomFrameIndexes(string customFrameFileName)
        {
            var indexes = customFrameFileName.Split(Constants.FileNameSeparator).Take(2).ToArray();
            var dokIndex = indexes[1].Substring(0, indexes[1].Length - 4);
            indexes[1] = dokIndex;
            return indexes;
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
                        case Constants.MATNR: UpdateAttribute(attrReference, indexes.First()); break;
                        case Constants.DOKNR: UpdateAttribute(attrReference, indexes.Last()); break;
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

            return 1 / Math.Truncate(mostOccuringTextValue / 2.5);
        }
    }
}
