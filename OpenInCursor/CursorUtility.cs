using System;
using System.Diagnostics;
using System.IO;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace OpenInCursor
{
    /// <summary>
    /// Utility class for common Cursor operations
    /// </summary>
    internal static class CursorUtility
    {
        /// <summary>
        /// Finds the Cursor executable path
        /// </summary>
        /// <returns>Path to cursor.exe if found, null otherwise</returns>
        public static string FindCursorExecutable()
        {
            string cursorExePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), 
                "Programs", "cursor", "cursor.exe");
            
            return File.Exists(cursorExePath) ? cursorExePath : null;
        }

        /// <summary>
        /// Opens a file or folder in Cursor
        /// </summary>
        /// <param name="package">The VS package for showing error messages</param>
        /// <param name="path">Path to open in Cursor</param>
        /// <returns>True if successful, false otherwise</returns>
        public static bool OpenInCursor(AsyncPackage package, string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                ShowErrorMessage(package, "No path provided.");
                return false;
            }

            // Validate path exists
            if (!File.Exists(path) && !Directory.Exists(path))
            {
                ShowErrorMessage(package, $"Path not found: {path}");
                return false;
            }

            string cursorPath = FindCursorExecutable();
            if (string.IsNullOrEmpty(cursorPath))
            {
                ShowWarningMessage(package, "Cursor not found. Please make sure Cursor is installed.");
                return false;
            }

            try
            {
                System.Diagnostics.Process.Start(new ProcessStartInfo
                {
                    FileName = cursorPath,
                    Arguments = $"\"{path}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                });
                return true;
            }
            catch (Exception ex)
            {
                ShowErrorMessage(package, $"Failed to open in Cursor: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Shows an error message box
        /// </summary>
        /// <param name="package">The VS package</param>
        /// <param name="message">Error message to display</param>
        public static void ShowErrorMessage(AsyncPackage package, string message)
        {
            VsShellUtilities.ShowMessageBox(
                package,
                message,
                "Open in Cursor",
                OLEMSGICON.OLEMSGICON_CRITICAL,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        /// <summary>
        /// Shows a warning message box
        /// </summary>
        /// <param name="package">The VS package</param>
        /// <param name="message">Warning message to display</param>
        public static void ShowWarningMessage(AsyncPackage package, string message)
        {
            VsShellUtilities.ShowMessageBox(
                package,
                message,
                "Open in Cursor",
                OLEMSGICON.OLEMSGICON_WARNING,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
} 