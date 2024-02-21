using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using DialogResult = System.Windows.Forms.DialogResult;
using OpenFileDialog = Autodesk.AutoCAD.Windows.OpenFileDialog;

namespace AcGraphicToFrame.Helpers
{
    internal static class FileManager
    {
        internal static string[] GetSelectedFileNames()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog("Select DWG files to copy from", null, "dwg", "DwgFileToLink", OpenFileDialog.OpenFileDialogFlags.AllowMultiple);
            DialogResult dialogResult = openFileDialog.ShowDialog();

            if (dialogResult != DialogResult.OK)
            {
                Active.Editor.WriteMessage("Error retrieving files. Try select files again.");
                return Array.Empty<string>();
            }

            return openFileDialog.GetFilenames();
        }

        internal static string GetPathToFrameFile(string formatName, string rootFolder, double scale) // TODO: manage scale value
        {
            var frameFileName = string.Concat(formatName, Constants.FileExtensionSeparator, Constants.DwgExtension);
            return Path.Combine(rootFolder, Constants.FramesFolderName, frameFileName);
        }

        internal static string GetResultFileName(string[] indexes)
        {
            return string.Concat(indexes[0], Constants.FileNameSeparator,
                indexes[1], indexes[2], Constants.FileNameSeparator, 
                Constants.FileNamePostfix, 
                Constants.FileExtensionSeparator, 
                Constants.DwgExtension);
        }
    }
}
