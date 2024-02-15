using AcGraphicToFrame.Forms;
using AcGraphicToFrame.Helpers;
using AcGraphicToFrame.Managers;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using System.Linq;
using DialogResult = System.Windows.Forms.DialogResult;
using OpenFileDialog = Autodesk.AutoCAD.Windows.OpenFileDialog;

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

            foreach(var file in selectedFilesNames)
            {
                DwgManager.CloneDrawingToFrame(file);
            }

            Info fileProcessorInfo = new Info();
            fileProcessorInfo.SetTextInfo($"DWG processed: {selectedFilesNames.Count()}");
            Application.ShowModalDialog(null, fileProcessorInfo, false);
        }
    }
}
