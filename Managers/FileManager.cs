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

        internal static string GetPathToFrameFile(string formatName, string rootFolder)
        {
            var frameFileName = string.Concat(formatName, Constants.FileExtensionSeparator, Constants.DwgExtension);
            return Path.Combine(rootFolder, Constants.FramesFolderName, frameFileName);
        }

        internal static string GetResultFileName(string rootFolder)
        {
            var fullNameWithPath = GetCustomFramePath(rootFolder);
            return Path.GetFileNameWithoutExtension(fullNameWithPath);
        }

        internal static string GetCustomFramePath(string rootFolder)
        {
            var customFrameFilePath = Path.Combine(rootFolder, Constants.CustomFramesFolder);
            return Directory.GetFiles(customFrameFilePath, $"*.{Constants.DwgExtension}", SearchOption.TopDirectoryOnly).FirstOrDefault();
        }

        internal static void MoveCustomFrameToUsedFolder(string customFrameUsed, string rootFolder)
        {
            var customFrameName = Path.GetFileNameWithoutExtension(customFrameUsed);
            var usedCustomFrameName = string.Concat(customFrameName, Constants.UsedFramePostfix, Constants.FileExtensionSeparator, Constants.DwgExtension);
            var usedCustomFramePath = Path.Combine(rootFolder, Constants.UsedFramesFolderName, usedCustomFrameName);
            File.Move(customFrameUsed, usedCustomFramePath);

        }
    }
}
