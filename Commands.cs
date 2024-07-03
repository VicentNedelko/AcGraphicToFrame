using AcGraphicToFrame.Exceptions;
using AcGraphicToFrame.Helpers;
using AcGraphicToFrame.Managers;
using Autodesk.AutoCAD.Runtime;
using OfficeOpenXml;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using MessageBox = System.Windows.MessageBox;

namespace AcGraphicToFrame
{
    public static class Commands
    {
        [CommandMethod("SELECT_FILES")]
        public static void SelectFiles()
        {
            var selectedFilesNames = FileManager.GetSelectedFileNames();
            if (selectedFilesNames.Count() == 0)
            {
                MessageBox.Show("DWG files list is empty.", "DWG file list ERROR", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);

                return;
            }

            var rootFolder = Path.GetDirectoryName(selectedFilesNames[0]);

            var promptDivider = Active.Editor.GetString("\nEnter divider value:");

            if (promptDivider.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
            {
                MessageBox.Show("Divder value fault.", "Divider value ERROR", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);

                return;
            }

            if (!Double.TryParse(promptDivider.StringResult, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var divider))
            {
                MessageBox.Show("Divider format error.", "Divider format ERROR", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);

                return;
            }

            try
            {
                var excelFileInfo = XlsxManager.GetExcelFileInfo(rootFolder);

                using (var excel = new ExcelPackage(excelFileInfo))
                {
                    var sheet = excel.Workbook.Worksheets[1];
                    int successfullyProcessedFilesNumber = default;
                    int totallyProcessedFilesNumber = default;
                    var progressMeter = new ProgressMeter();
                    var progressLimit = selectedFilesNames.Count();
                    progressMeter.SetLimit(progressLimit);
                    progressMeter.Start("DWG files converter");

                    for (var i = 0; i < selectedFilesNames.Count(); i++)
                    {
                        string status = string.Empty;
                        var fileName = string.Empty;
                        progressMeter.MeterProgress();

                        try
                        {
                            (status, fileName) = DwgManager.CloneDrawingToFrame(selectedFilesNames[i], sheet, divider);
                            XlsxManager.FillResult(sheet, status, fileName);
                        }
                        catch (MissedCustomInfoException)
                        {
                            break;
                        }

                        if (status == "SUCCESS")
                        {
                            successfullyProcessedFilesNumber++;
                        }
                        totallyProcessedFilesNumber++;

                        var messageOnEditor = ((double)totallyProcessedFilesNumber / selectedFilesNames.Count()).ToString("###.# %");
                        Active.Editor.WriteMessage($"Processed - {messageOnEditor}\n");
                        System.Windows.Forms.Application.DoEvents();
                    }

                    progressMeter.Stop();

                    excel.Save();

                    if (successfullyProcessedFilesNumber > 0)
                    {
                        MessageBox.Show($"Successfully processed DWG files: {successfullyProcessedFilesNumber}.\nCheck Protocol.xlsx for more information.", "DWG processor", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (FileNotFoundException)
            {
                return;
            }
        }
    }
}
