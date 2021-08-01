//-----------------------------------------------------------------------
// <copyright file="Widget.cs" company="Sprocket Enterprises">
//     Copyright (c) Sprocket Enterprises. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
namespace LKZ.SSMSUtils
{
    using EnvDTE;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;
    using System;
    using System.Collections.Specialized;
    using System.ComponentModel.Design;
    using System.Globalization;
    using System.Threading;
    using System.Threading.Tasks;
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
    /// 
    /// Configuration:
    ///   C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\devenv.exe
    ///   /rootsuffix Exp
    /// 
    ///   C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\ssms.exe
    ///   /log
    ///   C:\Program Files (x86)\Microsoft SQL Server Management Studio 18\Common7\IDE\Extensions\LKZ.SSMSUtils
    ///   %AppData%\Microsoft\AppEnv\15.0\ActivityLog.xml
    /// </summary>
    internal sealed class MacrossExample
    {
        StringDictionary myStringDictionary = new StringDictionary()
        {
            { "red", "#ff0000" },
            { "green", "#00cc00"},
            { "blue", "#0000ff"},
        };

        private DTE2 dte;
        private TextDocument textDocument;
        private EnvDTE80.Events2 dteEvents;
        private TextDocumentKeyPressEvents keyPressEvents;

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
        /// Initializes a new instance of the <see cref="MacrossExample"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private MacrossExample(AsyncPackage package, OleMenuCommandService commandService)
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
        public static MacrossExample Instance
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
            // Switch to the main thread - the call to AddCommand in Macross's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new MacrossExample(package, commandService);
            Instance.dte = (DTE2)(await package.GetServiceAsync(typeof(DTE)));
            Instance.textDocument = (TextDocument)Instance.dte.ActiveDocument.Object("TextDocument");
            Instance.dteEvents = (EnvDTE80.Events2)Instance.dte.Events;
            Instance.keyPressEvents = Instance.dteEvents.TextDocumentKeyPressEvents[Instance.textDocument];
            Instance.keyPressEvents.BeforeKeyPress += Instance.TextDocument_KeyPress;

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
            ThreadHelper.ThrowIfNotOnUIThread();

            // TODO: Abrir un popup con la configuracion de mis macros.

            //string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            //string title = "Macross";

            //Task.Run(async () => await DoStuff());

            //VsShellUtilities.ShowMessageBox(
            //    this.package,
            //    message,
            //    title,
            //    OLEMSGICON.OLEMSGICON_INFO,
            //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        //private async Task DoStuff()
        //{
        //    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        //    DTE2 dte = (DTE2)(await ServiceProvider.GetServiceAsync(typeof(DTE)));
        //    TextSelection ts = (TextSelection)dte.ActiveDocument.Selection;
        //    ProjectItem prj = (ProjectItem)dte.ActiveDocument.ProjectItem;
        //    string t = ts.Text;

        //    TextDocument textDocument = (TextDocument)dte.ActiveDocument.Object("TextDocument");
        //    string queryText = textDocument.Selection.Text;
        //    if (string.IsNullOrEmpty(queryText))
        //    {
        //        EditPoint startPoint = textDocument.StartPoint.CreateEditPoint();
        //        queryText = startPoint.GetText(textDocument.EndPoint);
        //    }

        //    Document document = (Document)dte.ActiveDocument.Object("Document");

        //    TextEditorEvents teev = dte.Events.TextEditorEvents[textDocument];
        //    teev.LineChanged += TextDocument_LineChanged;

        //    EnvDTE80.Events2 events = (EnvDTE80.Events2)dte.Events;
        //    TextDocumentKeyPressEvents tdkpev = events.TextDocumentKeyPressEvents[textDocument];
        //    tdkpev.BeforeKeyPress += TextDocument_KeyPress;

        //    EditPoint objEP = textDocument.StartPoint.CreateEditPoint();
        //    objEP.Insert("SELECT TOP 10 * FROM [].[]");
        //}

        private void TextDocument_KeyPress(string Keypress, TextSelection Selection, bool InStatementCompletion, ref bool CancelKeypress)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if ((Keypress == " ") || (Keypress == "\t"))
            {
                EditPoint ep = Selection.ActivePoint.CreateEditPoint();
                EditPoint sp = ep.CreateEditPoint();
                sp.CharLeft(1);
                while (true)
                {
                    string txt = sp.GetText(ep);
                    if (myStringDictionary.ContainsKey(txt))
                    {
                        sp.Delete(txt.Length);
                        sp.Insert(myStringDictionary[txt]);
                        CancelKeypress = true;
                        return;
                    }

                    sp.CharLeft(1);
                    if ((ep.Line != sp.Line) || ((ep.DisplayColumn == 1) && (ep.Line == 1)))
                        break;
                }
            }
        }
    }
}
