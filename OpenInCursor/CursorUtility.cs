using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Diagnostics;
using System.IO;

namespace OpenInCursor
{
    /// <summary>
    /// Utility class for common Cursor operations
    /// </summary>
    internal static class CursorUtility
    {
        // Cache for the Cursor executable path
        private static string _cachedCursorPath;
        private static bool _hasSearched = false;

        /// <summary>
        /// Finds the Cursor executable path from PATH environment variable.
        /// Uses lazy initialization - called automatically on first use.
        /// </summary>
        private static void FindCursorExecutable()
        {
            // Return if already searched
            if (_hasSearched)
            {
                return;
            }

            // Search in PATH environment variable
            string pathEnv = Environment.GetEnvironmentVariable("PATH");
            if (!string.IsNullOrEmpty(pathEnv))
            {
                var pathDirectories = pathEnv.Split(Path.PathSeparator);

                // Look for directories containing "cursor" in their path
                foreach (var directory in pathDirectories)
                {
                    if (string.IsNullOrWhiteSpace(directory))
                        continue;

                    try
                    {
                        // Check if the directory path contains "cursor" (case-insensitive)
                        if (directory.IndexOf("cursor", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            string cursorExePath = Path.Combine(directory, "cursor.cmd");
                            if (File.Exists(cursorExePath))
                            {
                                _cachedCursorPath = cursorExePath;
                                _hasSearched = true;
                                return;
                            }
                        }
                    }
                    catch
                    {
                        // Ignore invalid paths or access errors
                        continue;
                    }
                }
            }

            // Mark as searched even if not found
            // Don't show error during initialization to avoid UI deadlock
            _hasSearched = true;
            _cachedCursorPath = null;
        }

        /// <summary>
        /// Gets the cached Cursor executable path.
        /// If not searched yet, performs the search on first call (lazy initialization).
        /// </summary>
        /// <returns>Path to cursor.cmd if found, null otherwise</returns>
        private static string GetCursorPath()
        {
            // Lazy initialization: search on first use
            if (!_hasSearched)
            {
                FindCursorExecutable();
            }
            
            return _cachedCursorPath;
        }

        /// <summary>
        /// Opens a file or folder in Cursor
        /// </summary>
        /// <param name="package">The VS package for showing error messages</param>
        /// <param name="path">Path to open in Cursor</param>
        /// <returns>True if successful, false otherwise</returns>
        public static bool OpenInCursor(AsyncPackage package, string path)
        {
            return OpenInCursor(package, path, -1, -1);
        }

        /// <summary>
        /// Opens a file in Cursor at specific line and column position
        /// </summary>
        /// <param name="package">The VS package for showing error messages</param>
        /// <param name="path">Path to open in Cursor</param>
        /// <param name="line">Line number (1-based), -1 to ignore</param>
        /// <param name="column">Column number (1-based), -1 to ignore</param>
        /// <returns>True if successful, false otherwise</returns>
        public static bool OpenInCursor(AsyncPackage package, string path, int line, int column)
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

            string cursorPath = GetCursorPath();
            if (string.IsNullOrEmpty(cursorPath))
            {
                ShowErrorMessage(package, 
                    "Cursor executable not found in PATH.\n\n" +
                    "Please make sure:\n" +
                    "1. Cursor is installed\n" +
                    "2. Cursor is added to your system PATH\n" +
                    "3. Visual Studio is restarted after adding to PATH");
                return false;
            }

            try
            {
                // Build launch arguments
                string arguments;
                if (line > 0 && column > 0)
                {
                    // Use goto parameter to navigate to specified line and column
                    arguments = $"-g \"{path}:{line}:{column}\"";
                }
                else if (line > 0)
                {
                    // Only navigate to specified line
                    arguments = $"-g \"{path}:{line}\"";
                }
                else
                {
                    // Open file normally
                    arguments = $"\"{path}\"";
                }

                System.Diagnostics.Process.Start(new ProcessStartInfo
                {
                    FileName = cursorPath,
                    Arguments = arguments,
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