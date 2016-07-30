//------------------------------------------------------------------------------
// <copyright file="CreateDeployTaskCommand.cs" company="Company">
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
using Microsoft.Build.Construction;
using Microsoft.VisualStudio;
using System.Reflection;
using System.IO;

namespace DeployMsbuildExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CreateDeployTaskCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 256;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("915dc4e4-1750-4024-8007-a41f42b88678");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;


        DTE _dte;
        DTE dte
        {
            get
            {
                if (_dte == null)
                    _dte = ServiceProvider.GetService(typeof(DTE)) as DTE;
                return _dte;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CreateDeployTaskCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private CreateDeployTaskCommand(Package package)
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
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static CreateDeployTaskCommand Instance
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
            Instance = new CreateDeployTaskCommand(package);
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
            EnvDTE.Project selectedProject = null;
            if (dte.SelectedItems.Count > 0)
                selectedProject = (dte.SelectedItems.Item(1)).Project;
            if (selectedProject == null) return;

            var assembly = Assembly.GetExecutingAssembly();
            using(var sw = new StreamWriter(Path.Combine(Path.GetDirectoryName(selectedProject.FullName), "Deploy.proj")))
            {
                using(var sr = new StreamReader(assembly.GetManifestResourceStream("DeployMsbuildExtension.Deploy.proj")))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }
                //selectedProject.ProjectItems.AddFromTemplate("Deploy.proj", "Deploy.proj");
            //selectedProject.ProjectItems.AddFromFile("Deploy.proj");

            var solution = ServiceProvider.GetService(typeof(SVsSolution)) as SVsSolution;
            var solutionProj = solution as IVsSolution4;
            var solutionTop = solution as IVsSolution;

            IVsHierarchy hierarchy;

            ErrorHandler.ThrowOnFailure(solutionTop.GetProjectOfUniqueName(selectedProject.FullName, out hierarchy));
            

            if (hierarchy != null)
            {
                Guid projectGuid;

                ErrorHandler.ThrowOnFailure(hierarchy.GetGuidProperty(
                            VSConstants.VSITEMID_ROOT,
                            (int)__VSHPROPID.VSHPROPID_ProjectIDGuid,
                            out projectGuid));

                //ErrorHandler.ThrowOnFailure(solutionProj.UnloadProject(projectGuid, 
                //    (uint)_VSProjectUnloadStatus.UNLOADSTATUS_UnloadedByUser));
    
                var root = ProjectRootElement.Open(selectedProject.FullName);
                root.AddImport("Deploy.proj");
                root.AddItem("None", "Deploy.proj");
                root.Save();

                ErrorHandler.ThrowOnFailure(solutionProj.ReloadProject(projectGuid));
            }

            
        }
    }
}
