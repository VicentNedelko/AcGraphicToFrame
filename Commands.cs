using AcGraphicToFrame.Forms;
using AcGraphicToFrame.Helpers;
using AcGraphicToFrame.Managers;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using OfficeOpenXml;
using System.IO;
using System.Linq;

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
                Info fileSelectorFormInfo = new Info();
                fileSelectorFormInfo.SetTextInfo("DWG file selection error.");
                Application.ShowModalDialog(null, fileSelectorFormInfo, false);

                return;
            }

            var rootFolder = Path.GetDirectoryName(selectedFilesNames[0]);

            var excelFileInfo = XlsxManager.GetExcelFileInfo(rootFolder);

            using (var excel = new ExcelPackage(excelFileInfo))
            {
                var sheet = excel.Workbook.Worksheets[1];
                for (var i = 0; i < selectedFilesNames.Count(); i++)
                {
                    var status = DwgManager.CloneDrawingToFrame(selectedFilesNames[i], sheet);
                    XlsxManager.FillResult(sheet, status);
                }

                excel.Save();

                Info fileProcessorInfo = new Info();
                fileProcessorInfo.SetTextInfo($"DWG processed: {selectedFilesNames.Count()}");
                Application.ShowModalDialog(null, fileProcessorInfo, false);
            }

        }
    }
}
