//-----------------------------------------------------------------------
// <copyright file="MacrossPackage.cs" company="LKZ">
//     Copyright (c) LKZ. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace LKZ.SSMSUtils
{
    using System;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Task = System.Threading.Tasks.Task;
    using EnvDTE;
    using EnvDTE80;
    using Microsoft.VisualStudio;
    using Microsoft.VisualStudio.Shell;
    using System.Windows.Forms;
    using Microsoft.VisualStudio.Shell.Interop;

    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NotBuildingAndNotDebugging_string, PackageAutoLoadFlags.BackgroundLoad)] // Auto-load for dynamic menu enabling/disabling; this context seems to work for SSMS and VS
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(MacrossPackage.PackageGuidString)]
    public sealed class MacrossPackage : AsyncPackage
    {
        /// <summary>
        /// MacrossPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "59213989-55cb-4f5e-b2f0-147254bbb8df";

        private const string macrosFilename = "Macross.txt";

        //private StringDictionary myStringDictionary = new StringDictionary()
        //{
        //    { "red", "#ff0000" },
        //    { "green", "#00cc00"},
        //    { "blue", "#0000ff"},
        //};

        private Dictionary<string, string> macros =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "LKZ", "lukiller:"}
            };

        private DTE2 dte;
        private TextDocument textDocument;
        private EnvDTE80.Events2 dteEvents;
        private TextDocumentKeyPressEvents keyPressEvents;

        /// <summary>
        /// Initializes a new instance of the <see cref="MacrossPackage"/> class.
        /// </summary>
        public MacrossPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            try
            {
                // When initialized asynchronously, the current thread may be a background thread at this point.
                // Do any initialization that requires the UI thread after switching to the UI thread.
                await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

                this.dte = (DTE2)(await this.GetServiceAsync(typeof(DTE)));
                this.dteEvents = (EnvDTE80.Events2)this.dte.Events;
                // dteEvents.DocumentEvents.DocumentOpened += DTE_DocumentOpened; // These Document events are not firing :_(
                // dteEvents.DocumentEvents.DocumentOpening += DTE_DocumentOpening;
                dteEvents.WindowEvents.WindowActivated += DTE_WindowActivated;

                await LKZ.SSMSUtils.MacrossCommand.InitializeAsync(this);
                this.LoadMacrosFromFile();
                base.Initialize();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"MacrossPackage.InitializeAsync(): {ex.Message}", @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Handler for the Activated event. It is triggered every time any 
        /// window got focus.
        /// </summary>
        /// <param name="gotFocus">The window that got focus.</param>
        /// <param name="lostFocus">The window that lost focus.</param>
        private void DTE_WindowActivated(Window gotFocus, Window lostFocus)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                if (!gotFocus.Kind.Equals("Document", StringComparison.OrdinalIgnoreCase)
                    || this.dte.ActiveDocument == null)
                {
                    return;
                }

                if (this.textDocument == null)
                {
                    this.textDocument = this.dte.ActiveDocument?.Object("TextDocument") as TextDocument;
                    if (this.textDocument == null)
                    {
                        return;
                    }

                    this.keyPressEvents = this.dteEvents.TextDocumentKeyPressEvents[this.textDocument];
                    this.keyPressEvents.BeforeKeyPress += this.TextDocument_KeyPress;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"MacrossPackage.DTE_WindowActivated(): {ex.Message}", @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Handler for the KeyPress event. Replaces the text with the macro.
        /// OLD: It changes the color to it's hexadecimal representation.
        /// </summary>
        /// <param name="keypress">The key pressed.</param>
        /// <param name="selection">The selected text.</param>
        /// <param name="inStatementCompletion">Whether the statement is in completion state.</param>
        /// <param name="cancelKeypress">Whether the event is being cancelled.</param>
        private void TextDocument_KeyPress(string keypress, TextSelection selection, bool inStatementCompletion, ref bool cancelKeypress)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                if (keypress == "\t")
                {
                    EditPoint ep = selection.ActivePoint.CreateEditPoint();
                    EditPoint sp = ep.CreateEditPoint();
                    sp.CharLeft(1);
                    while (true)
                    {
                        string txt = sp.GetText(ep);
                        if (macros.ContainsKey(txt))
                        {
                            // verificar si empieza con espacio o es la primera columna
                            sp.Delete(txt.Length);
                            sp.Insert(macros[txt]);
                            cancelKeypress = true;
                            return;
                        }

                        sp.CharLeft(1);
                        if ((ep.Line != sp.Line) || ((ep.DisplayColumn == 1) && (ep.Line == 1)))
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"MacrossPackage.TextDocument_KeyPress(): {ex.Message}", @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Loads the macros from a CSV file.
        /// </summary>
        private void LoadMacrosFromFile()
        {
            try
            {
                string line;
                var file = new StreamReader(macrosFilename);
                while ((line = file.ReadLine()) != null)
                {
                    var macro = line.Split(',');
                    if (macro.Length == 2)
                    {
                        this.macros.Add(macro[0], macro[1]);
                    }
                }

                file.Close();
                SetStatus("Macross loaded successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"MacrossPackage.LoadMacrosFromFile(): {ex.Message}", @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Displays a message in the status bar.
        /// </summary>
        /// <param name="message">The message.</param>
        public void SetStatus(string message)
        {
            var statusBar = GetService(typeof(SVsStatusbar)) as IVsStatusbar;
            if (statusBar != null)
            {
                int frozen;
                statusBar.IsFrozen(out frozen);
                if (!Convert.ToBoolean(frozen))
                {
                    if (message == null)
                    {
                        statusBar.Clear();
                    }
                    else
                    {
                        statusBar.SetText(message);
                    }
                }
            }
        }
    }
}
