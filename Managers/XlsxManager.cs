using OfficeOpenXml;
using System.Drawing;
using System.IO;
using System.Linq;

namespace AcGraphicToFrame.Managers
{
    internal static class XlsxManager
    {
        internal static void FillResult(ExcelWorksheet sheet, string status)
        {
            var freeRowIndex = GetFreeRowIndex(sheet);
            sheet.Cells[$"X{freeRowIndex}"].Value = status;
            if(status != "SUCCESS")
            {
                sheet.Cells[$"X{freeRowIndex}"].Style.Fill.BackgroundColor.SetColor(Color.Red);
            }
        }

        internal static FileInfo GetExcelFileInfo(string rootFolder)
        {
            var excelFileName = string.Concat(Constants.ExcelFileName, Constants.FileExtensionSeparator, Constants.XlsxExtension);
            var excelFilePath = Path.Combine(rootFolder, excelFileName);
            return new FileInfo(excelFilePath);
        }

        internal static int GetFreeRowIndex(ExcelWorksheet sheet)
        {
            try
            {
                var firstFreeCell = sheet.Cells["X:X"]
                    .Where(x => string.IsNullOrEmpty(x.Value as string))
                    .Select(x => x)
                    .FirstOrDefault();

                return firstFreeCell.Start.Row;
            }
            catch
            {
                throw;
            }
        }

        
    }
}
