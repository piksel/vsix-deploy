//------------------------------------------------------------------------------
// <copyright file="DeployCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using Microsoft.Build.Evaluation;
using Microsoft.Build.Logging;
using Microsoft.Build.Framework;
using Microsoft.VisualStudio;
using System.Runtime.InteropServices;

namespace DeployMsbuildExtension
{
    using Project = Microsoft.Build.Evaluation.Project;

    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class DeployCommand
    {

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("3e21ecd8-d843-4c95-b6de-2f623b5793fa");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        private const string paneTitle = "Deploy";
        private Guid paneGuid = new Guid("C45CFDA1-A618-491F-9E58-66F8AF70AD60");

        DTE _dte;
        DTE dte
        {
            get
            {
                if(_dte == null)
                    _dte = ServiceProvider.GetService(typeof(DTE)) as DTE;
                return _dte;
            }
        }

        IVsOutputWindowPane _pane;
        private IVsOutputWindowPane Pane
        {
            get
            {
                if (_pane == null)
                {
                    IVsOutputWindow outWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;
                    outWindow.CreatePane(ref paneGuid, paneTitle, 1, 1);
                    outWindow.GetPane(ref paneGuid, out _pane);

                    dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput).Visible = true;
                }
                return _pane;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DeployCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private DeployCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new OleMenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);

                menuCommandID = new CommandID(CommandSet, CommandId + 0x0010);
                menuItem = new OleMenuCommand(this.MenuItemCallback, menuCommandID);
                menuItem.BeforeQueryStatus += MenuItem_BeforeQueryStatus;
                commandService.AddCommand(menuItem);
            }
        }

        private void MenuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            var item = (OleMenuCommand)sender;
            item.Text = "Deploy Selection";
            item.Visible = false;
            foreach (SelectedItem selectedItem in dte.SelectedItems)
            {
                if (selectedItem.Project != null)
                {
                    var project = selectedItem.Project;
                    item.Text = "Deploy " + project.Name;
                    item.Visible = true;
                    return;
                }
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static DeployCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
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
        public static void Initialize(Package package)
        {
            Instance = new DeployCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {

            string message = "An unknown error occured.";
            bool result = false;
            try
            {
                Pane.Activate();
                Pane.OutputString("Attempting to run Deploy task...\n");

                var project = getCurrentProject();
                if (project == null) throw new Exception("Cannot find target project!"); 
                var projectInstance = project.CreateProjectInstance();
                var logger = new OutputLogger(Pane);


                result = projectInstance.Build("Deploy", new[] { logger });

                message = "Deploy build ran " + (result ? "" : "un") + "successfully";


            }
            catch (Exception x)
            {
                message = "Encountered the following exception when trying to run the Deploy task:\n" + x.Message;
            }
            finally
            {
                Pane.OutputString("\n"+message);
            }
        }

        private T getGlobalService<T>() where T: class
        {
            return Package.GetGlobalService(typeof(T)) as T;
        }

        private Project getCurrentProject()
        {
            Project project;

            if (dte.ActiveDocument != null && dte.ActiveDocument.ProjectItem != null)
            {
                var projName = dte.ActiveDocument.ProjectItem.ContainingProject.FullName;
                project = ProjectCollection.GlobalProjectCollection.LoadProject(projName);
            }
            else
            {
                IntPtr hierarchyPointer = IntPtr.Zero;
                IntPtr selectionContainerPointer = IntPtr.Zero;
                try
                {
                    Object selectedObject = null;
                    IVsMultiItemSelect multiItemSelect;
                    uint projectItemId;

                    var monitorSelection = getGlobalService<SVsShellMonitorSelection>() as IVsMonitorSelection;

                    monitorSelection.GetCurrentSelection(out hierarchyPointer,
                                                         out projectItemId,
                                                         out multiItemSelect,
                                                         out selectionContainerPointer);

                    IVsHierarchy selectedHierarchy = Marshal.GetTypedObjectForIUnknown(
                                                         hierarchyPointer,
                                                         typeof(IVsHierarchy)) as IVsHierarchy;

                    if (selectedHierarchy != null)
                    {
                        ErrorHandler.ThrowOnFailure(selectedHierarchy.GetProperty(
                                                          projectItemId,
                                                          (int)__VSHPROPID.VSHPROPID_ExtObject,
                                                          out selectedObject));
                    }

                    project = selectedObject as Project;

                    if (project == null)
                    {
                        //var activeProjects = dte.ActiveSolutionProjects as Project[];
                        //if (activeProjects != null && activeProjects.Length > 0)
                        //    project = activeProjects[0];
                        foreach(SelectedItem si in dte.SelectedItems)
                        {
                            var projName = si.Project.FullName;
                            project = ProjectCollection.GlobalProjectCollection.LoadProject(projName);
                            if (project != null) break;
                        }
                    }
                      //  project = dte.ActiveSolutionProjects.Projects.Item(projectItemId) as Project;
                }
                finally
                {
                    if (hierarchyPointer == IntPtr.Zero) Marshal.Release(hierarchyPointer);
                    if (selectionContainerPointer == IntPtr.Zero) Marshal.Release(selectionContainerPointer);
                }


            }

            return project;

        }
    }
}
