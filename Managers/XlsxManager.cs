using AcGraphicToFrame.Exceptions;
using OfficeOpenXml;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows;

namespace AcGraphicToFrame.Managers
{
    internal static class XlsxManager
    {
        internal static void FillResult(ExcelWorksheet sheet, string status, string filePath)
        {
            var freeRowIndex = GetFreeRowIndex(sheet);
            sheet.Cells[$"Y{freeRowIndex}"].Value = Path.GetFileName(filePath);
            sheet.Cells[$"X{freeRowIndex}"].Value = status;

            if (status != "SUCCESS")
            {
                sheet.Cells[$"X{freeRowIndex}"].Style.Fill.BackgroundColor.SetColor(Color.Red);
                return;
            }
            sheet.Cells[$"X{freeRowIndex}"].Style.Fill.BackgroundColor.SetColor(Color.Green);
        }

        internal static FileInfo GetExcelFileInfo(string rootFolder)
        {
            var excelFileName = string.Concat(Constants.ExcelFileName, Constants.FileExtensionSeparator, Constants.XlsxExtension);
            var excelFilePath = Path.Combine(rootFolder, excelFileName);
            if (File.Exists(excelFilePath))
            {
                return new FileInfo(excelFilePath);
            }
            MessageBox.Show("Protocol.xlsx file was not provided.", "Protocol.xlsx ERROR", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.Yes);
            throw new FileNotFoundException(excelFilePath);
        }

        internal static int GetFreeRowIndex(ExcelWorksheet sheet)
        {
            try
            {
                var firstFreeCell = sheet.Cells["X:X"]
                    .Where(x => string.IsNullOrEmpty(x.Value as string))
                    .Select(x => x)
                    .First();

                return firstFreeCell.Start.Row;
            }
            catch
            {
                throw new MissedCustomInfoException("No valid custom info provided in Protocol. Check Protocol.xlsx."); ;
            }
        }

        internal static string[] GetIndexesTable(ExcelWorksheet sheet, int freeRowIndex)
        {
            try
            {
                var materialN = Math.Truncate((double)sheet.Cells[$"B{freeRowIndex}"].Value).ToString(); // TODO: manage value if not double
                var documentN = Math.Truncate((double)sheet.Cells[$"C{freeRowIndex}"].Value).ToString();
                var sheetN = sheet.Cells[$"D{freeRowIndex}"].Value.ToString();

                return new[] { materialN, documentN, sheetN };
            }
            catch
            {
                throw new AttributesNotFoundException($"Block attribute(s) not found or have wrong format at {freeRowIndex} row.");
            }
        }

    }
}
