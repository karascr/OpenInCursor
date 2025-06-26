using System;
using System.ComponentModel.Design;
using System.IO;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;
using Task = System.Threading.Tasks.Task;
using Microsoft.VisualStudio;
using EnvDTE;
using EnvDTE80;

namespace OpenInCursor
{
    /// <summary>
    /// Solution Explorer context menu command handler
    /// </summary>
    internal sealed class SolutionExplorerCommand
    {
        /// <summary>
        /// Command IDs for Solution Explorer commands
        /// </summary>
        public const int OpenInCursorFileId = 0x0101;
        public const int OpenInCursorFolderId = 0x0102;
        public const int OpenInCursorProjectId = 0x0103;
        public const int OpenInCursorSolutionId = 0x0104;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("a7fd7efd-fb26-4ae2-b888-221254ed76c4");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="SolutionExplorerCommand"/> class.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private SolutionExplorerCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            // Add command handlers for each Solution Explorer context menu
            AddCommand(commandService, OpenInCursorFileId, ExecuteFileCommand);
            AddCommand(commandService, OpenInCursorFolderId, ExecuteFolderCommand);
            AddCommand(commandService, OpenInCursorProjectId, ExecuteProjectCommand);
            AddCommand(commandService, OpenInCursorSolutionId, ExecuteSolutionCommand);
        }

        private void AddCommand(OleMenuCommandService commandService, int commandId, EventHandler handler)
        {
            var menuCommandID = new CommandID(CommandSet, commandId);
            var menuItem = new MenuCommand(handler, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static SolutionExplorerCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                Instance = new SolutionExplorerCommand(package, commandService);
            }
        }

        private void ExecuteFileCommand(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _ = ExecuteFileCommandAsync().ConfigureAwait(false);
        }

        private void ExecuteFolderCommand(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _ = ExecuteFolderCommandAsync().ConfigureAwait(false);
        }

        private void ExecuteProjectCommand(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _ = ExecuteProjectCommandAsync().ConfigureAwait(false);
        }

        private void ExecuteSolutionCommand(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _ = ExecuteSolutionCommandAsync().ConfigureAwait(false);
        }



        private string GetSelectedItemPath()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            
            try
            {
                var dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
                if (dte?.SelectedItems != null)
                {
                    foreach (SelectedItem selectedItem in dte.SelectedItems)
                    {
                        if (selectedItem.ProjectItem != null)
                        {
                            // File or folder item
                            if (selectedItem.ProjectItem.Properties != null)
                            {
                                try
                                {
                                    var fullPath = selectedItem.ProjectItem.Properties.Item("FullPath")?.Value?.ToString();
                                    if (!string.IsNullOrEmpty(fullPath))
                                    {
                                        return fullPath;
                                    }
                                }
                                catch
                                {
                                    // If FullPath is not available, try to get the file name
                                    var fileName = selectedItem.ProjectItem.FileNames[1];
                                    if (!string.IsNullOrEmpty(fileName))
                                    {
                                        return fileName;
                                    }
                                }
                            }
                        }
                        else if (selectedItem.Project != null)
                        {
                            // Project item
                            return Path.GetDirectoryName(selectedItem.Project.FullName);
                        }
                    }
                }

                // Fallback: try to get solution folder
                if (dte?.Solution != null && !string.IsNullOrEmpty(dte.Solution.FullName))
                {
                    return Path.GetDirectoryName(dte.Solution.FullName);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting selected item path: {ex.Message}");
            }

            return null;
        }

        private async Task ExecuteFileCommandAsync()
        {
            await ExecuteCommandAsync(true);
        }

        private async Task ExecuteFolderCommandAsync()
        {
            await ExecuteCommandAsync(false);
        }

        private async Task ExecuteProjectCommandAsync()
        {
            await ExecuteCommandAsync(false);
        }

        private async Task ExecuteSolutionCommandAsync()
        {
            await ExecuteCommandAsync(false);
        }

        private async Task ExecuteCommandAsync(bool isFile)
        {
            try
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                string selectedPath = GetSelectedItemPath();
                if (string.IsNullOrEmpty(selectedPath))
                {
                    CursorUtility.ShowWarningMessage(this.package, "No item selected or unable to determine the path.");
                    return;
                }

                // For files, open the file directly
                // For folders/projects/solutions, open the containing folder
                string pathToOpen = selectedPath;
                if (!isFile)
                {
                    // For folders, projects, and solutions, open the directory
                    if (File.Exists(selectedPath))
                    {
                        pathToOpen = Path.GetDirectoryName(selectedPath);
                    }
                }

                CursorUtility.OpenInCursor(this.package, pathToOpen);
            }
            catch (Exception ex)
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                CursorUtility.ShowErrorMessage(this.package, $"An error occurred: {ex.Message}");
            }
        }
    }
} 