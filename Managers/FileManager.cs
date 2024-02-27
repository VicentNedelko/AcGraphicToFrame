using System;
using System.IO;
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

        internal static string GetPathToFrameFile(string formatName, string rootFolder, double scale)
        {
            var frameFileName = string.Concat(formatName, Constants.FileExtensionSeparator, Constants.DwgExtension);
            string formatFolderName = string.Empty;
            switch (scale)
            {
                case 10:
                    formatFolderName = "1_10";
                    break;
                case 20:
                    formatFolderName = "1_20";
                    break;
            };

            return Path.Combine(rootFolder, Constants.FramesFolderName, formatFolderName, frameFileName);
        }

        internal static string GetResultFileName(string[] indexes)
        {
            return string.Concat(indexes[0], 
                Constants.FileNameSeparator,
                indexes[1], indexes[2], 
                Constants.FileNameSeparator, 
                Constants.FileNamePostfix, 
                Constants.FileExtensionSeparator, 
                Constants.DwgExtension);
        }
    }
}
