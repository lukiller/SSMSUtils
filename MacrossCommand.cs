//-----------------------------------------------------------------------
// <copyright file="MacrossCommand.cs" company="LKZ">
//     Copyright (c) LKZ. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace LKZ.SSMSUtils
{
    using EnvDTE;
    using EnvDTE80;
    using LKZ.SSMSUtils.Views;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using System;
    using System.Collections.Specialized;
    using System.ComponentModel.Design;
    using System.Globalization;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using Task = System.Threading.Tasks.Task;

    /// <summary>
    /// Command handler
    /// 
    /// Doc:
    ///   https://stackoverflow.com/questions/55661806/how-to-create-an-extension-for-ssms-2019-v18
    ///   https://www.codeproject.com/Articles/5306973/Create-a-SQL-Server-Management-Studio-Extension
    ///   https://docs.microsoft.com/en-us/dotnet/api/envdte80.dte2?view=visualstudiosdk-2019
    /// 
    /// Examples:
    ///   https://www.ssmsboost.com/social/posts/t75-Feature-Request--Execute-current-statement
    ///   https://docs.microsoft.com/en-us/dotnet/api/envdte80.textdocumentkeypressevents?view=visualstudiosdk-2019
    ///   https://www.codeproject.com/Articles/1073839/Create-SQL-Server-Management-Studio-Addin (old, but examples with menu options in Object Explorer)
    /// 
    /// Configuration:
    ///   C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\devenv.exe
    ///   /rootsuffix Exp
    /// 
    ///   C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\ssms.exe
    ///   /log
    ///   C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\Extensions\LKZ.SSMSUtils
    ///   %AppData%\Microsoft\AppEnv\15.0\ActivityLog.xml
    ///   
    /// If you want your extension to work with sql related features then you are going to need to add references 
    /// to some SSMS dll files: 
    ///  - SqlWorkbench.Interfaces.dll located in the ManagementStudio folder.
    ///  - SqlPackageBase.dll contains some GUIDs that will be useful for the package ProvideAutoLoad attribute.
    /// </summary>
    internal sealed class MacrossCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("4ab3d463-5dcc-4a4f-aabe-1f7bda01a103");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="MacrossCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private MacrossCommand(AsyncPackage package, OleMenuCommandService commandService)
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
        public static MacrossCommand Instance
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
            try
            {
                // Switch to the main thread - the call to AddCommand in Macross's constructor requires
                // the UI thread.
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

                OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
                Instance = new MacrossCommand(package, commandService);
                //Instance.dte = (DTE2)(await package.GetServiceAsync(typeof(DTE)));
                //Instance.dteEvents = (EnvDTE80.Events2)Instance.dte.Events;
                //Instance.dteEvents.DocumentEvents.DocumentOpened += Instance.DocumentEvents_DocumentOpened;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"MacrossCommand.InitializeAsync(): {ex.Message}", @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                var dialog = new MacrossView();
                dialog.ShowDialog();
                //MessageBox.Show($"Hola mundo cruel!", @"Macross", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // TODO: Abrir un popup con la configuracion de mis macros o mandar a abrir el TXT con el notepad.
            }
            catch (Exception ex)
            {
                MessageBox.Show($"MacrossCommand.Execute(): {ex.Message}", @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
