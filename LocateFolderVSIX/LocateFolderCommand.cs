using EnvDTE;
using EnvDTE80;
using LocateFolderVSIX.Utilities;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace LocateFolderVSIX
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class LocateFolderCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("bc98d2f7-f2af-425a-b063-87a2a69e725a");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;
        


        /// <summary>
        /// Initializes a new instance of the <see cref="LocateFolderCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private LocateFolderCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));
            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static LocateFolderCommand Instance
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
            // Switch to the main thread - the call to AddCommand in LocateFolderCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new LocateFolderCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private async void Execute(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            //ThreadHelper.ThrowIfNotOnUIThread();
            //string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            //string title = "LocateFolderCommand";

            //// Show a message box to prove we were here
            //VsShellUtilities.ShowMessageBox(
            //    this.package,
            //    message,
            //    title,
            //    OLEMSGICON.OLEMSGICON_INFO,
            //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);


            try
            {

                //var task = (Task)((UIHierarchy)((DTE2)this.ServiceProvider.GetServiceAsync(typeof(DTE)))
                //    .Windows.Item("{3AE79031-E1BC-11D0-8F78-00A0C9110057}").Object);
                //var selectedItems = task.GetType().GetProperty("Result").GetValue(task) as Object[];
                var dte = await package.GetServiceAsync(typeof(DTE)).ConfigureAwait(false) as DTE2;
               
                var selectedItems = ((UIHierarchy)(dte.Windows.Item("{3AE79031-E1BC-11D0-8F78-00A0C9110057}").Object)).SelectedItems as object[];
                   
                if (selectedItems != null)
                {
                    LocateFile.FilesOrFolders((IEnumerable<string>)(from t in selectedItems
                                                                    where (t as UIHierarchyItem)?
                                                                    .Object is ProjectItem
                                                                    select ((ProjectItem)
                                                                    ((UIHierarchyItem)t).Object).
                                                                    FileNames[1]));
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("MyMessage: " + ex.Message);
            }
            
        }
    }
}
