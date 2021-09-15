using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace ExtractInterfaceBulk
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ExtractInterfaceBulk
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("6417f7f9-22d8-4fff-8093-01cd6e48b34d");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtractInterfaceBulk"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ExtractInterfaceBulk(AsyncPackage package, OleMenuCommandService commandService)
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
        public static ExtractInterfaceBulk Instance
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
            // Switch to the main thread - the call to AddCommand in ExtractInterfaceBulk's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ExtractInterfaceBulk(package, commandService);
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
            WriteLog("**************************************************");
            WriteLog("Start to execute of this extension.");


            var ignoreFolders = new string[] { "bin", "obj", "Interfaces", "DummyEntity", "Extensions", "" };
            var ignoreFiles = new string[] { "AssemblyInfo.cs", };
            var interfaceInherit = "IAdoEntity";

            ThreadHelper.ThrowIfNotOnUIThread();
            string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            string title = "Extract Interface Bulk";

            ThreadHelper.ThrowIfNotOnUIThread();
            var vsDte2 = Package.GetGlobalService(typeof(DTE)) as DTE2;
            var solution = vsDte2.Solution;

            // Get Solution and Project paths
            var projects = (UIHierarchyItem[])vsDte2?.ToolWindows.SolutionExplorer.SelectedItems;

            var project = projects[0].Object as Project;

            var solutionName = Path.GetFileName(solution.FullName);
            var solutionDir = Path.GetDirectoryName(solution.FullName);
            var projectName = Path.GetFileName(project.FullName);
            var projectDir = Path.GetDirectoryName(project.FullName);

            var needAddClasses = new List<string>();
            var alreadyHasInterfaceClasses = new List<string>();
            var allFiles = FindSubordinaryFiles(projectDir, ignoreFolders);
            allFiles.ForEach(file =>
            {
                var processedClass = ExtractInterface(file, vsDte2, alreadyHasInterfaceClasses, interfaceInherit);
                if (!string.IsNullOrEmpty(processedClass))
                    needAddClasses.Add(processedClass);
            });

            WriteLog("Writing final summary message.");
            var summaryMsg = string.Format(@"Processed Classes: {0} \r\n Already has Interface: {1}", string.Join(",", needAddClasses.ToArray()), string.Join(",", alreadyHasInterfaceClasses.ToArray()));
            WriteLog(summaryMsg);

            System.Windows.Forms.SendKeys.Flush();
            // Show message box about Processed Classes
            VsShellUtilities.ShowMessageBox(
                this.package,
                summaryMsg,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        [DllImport("user32.dll", EntryPoint = "keybd_event", SetLastError = true)]
        public static extern void keybd_event(Keys bVk, byte bScan, uint dwFlags, uint dwExtraInfo);
        public const int KEYEVENTF_KEYUP = 2;

        /// <summary>
        /// return processed class name
        /// </summary>
        private string ExtractInterface(string file, DTE2 dte2, List<string> alreadyHasInterfaceClasses, string interfaceInherit)
        {
            var logMsg = string.Empty;

            var processedClass = string.Empty;
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            // only process class files, and .cs is first judgement
            if (file.EndsWith(".cs"))
            {
                Debug.WriteLine("begin process file: " + file);
                logMsg += "\n" + "begin process file: " + file;
                WriteLog("begin process file: " + file);

                // Open class file
                Debug.WriteLine("before open file: " + file);
                logMsg += "\n" + "before open file: " + file;
                WriteLog("before open file: " + file);

                var openFileWindow = dte2.ItemOperations.OpenFile(file);
                openFileWindow.Activate();
                //dte2.Documents.Open(file);

                System.Threading.Thread.Sleep(100);
                //while (!dte2.ItemOperations.IsFileOpen(file))
                //{
                //    System.Threading.Thread.Sleep(1000);
                //    Debug.WriteLine("opening file: " + file);
                //    logMsg += "\n" + "opening file: " + file;
                //    Debug.WriteLine("opening file: " + file);
                //}
                //Debug.WriteLine("opened file: " + file);
                //logMsg += "\n" + "opened file: " + file;
                WriteLog("opened file: " + file);

                var iWpfTextViewHost = GetCurrentViewHost();
                //Keep below, this may be used in future.
                //var iTextDocument = GetTextDocumentForView(iWpfTextViewHost);
                //var xx = iTextDocument.TextBuffer;

                var view = iWpfTextViewHost.TextView;
                var textToSearch = "public class";
                var indexOfPublicClass = view.TextSnapshot.GetText().IndexOf(textToSearch);
                if (indexOfPublicClass < 0) return processedClass;
                view.Caret.MoveTo(view.GetTextViewLineContainingBufferPosition(view.TextSnapshot.GetLineFromPosition(0).Start));
                view.Caret.MoveTo(view.GetTextViewLineContainingBufferPosition(view.TextSnapshot.GetLineFromPosition(indexOfPublicClass).Start));

                openFileWindow.Activate();

                var allText = view.TextSnapshot.GetText();
                Regex regex = new Regex(@"public class (\S+)");
                if (regex.IsMatch(allText))
                {
                    processedClass = regex.Match(allText).Value.Substring(13);
                }
                else
                {
                    Debug.WriteLine("find no class: " + processedClass);
                    logMsg += "\n" + "find no class: " + processedClass;
                    WriteLog("find no class: " + processedClass);

                    return processedClass;
                }

                Debug.WriteLine("find class: " + processedClass);
                logMsg += "\n" + "find class: " + processedClass;
                WriteLog("find class: " + processedClass);

                var interfaceName = string.Format("I{0}", processedClass);
                var interfacePart1 = string.Format(": I{0}", processedClass);
                var interfacePart2 = string.Format(", I{0}", processedClass);
                if (allText.Contains(interfacePart1) || allText.Contains(interfacePart2))
                {
                    WriteLog("already extract interface: " + processedClass);
                    alreadyHasInterfaceClasses.Add(processedClass);
                    return string.Empty;
                }

                openFileWindow.Activate();
                SimulateKeysToExtractInterface();
                openFileWindow.Activate();

                string fileName = System.IO.Path.GetFileName(file);
                string interfaceFile = file.Replace(processedClass + ".cs", "I" + processedClass + ".cs");
                if (!File.Exists(interfaceFile))
                {
                    WriteLog("interface file not found!");
                    return string.Empty;
                }

                //var openFileWindow2 = dte2.ItemOperations.OpenFile(interfaceFile);
                //openFileWindow2.Activate();
                //ModifyInterface(interfaceName, openFileWindow2);

                Debug.WriteLine("extract interface successfully: " + processedClass);
                logMsg += "\n" + "extract interface successfully: " + processedClass;
                WriteLog("extract interface successfully: " + processedClass);
            }

            return processedClass;
        }

        private void ModifyInterface(string interfaceName, Window openFileWindow2)
        {
            return;

            var iWpfTextViewHost = GetCurrentViewHost();
            //Keep below, this may be used in future.
            //var iTextDocument = GetTextDocumentForView(iWpfTextViewHost);
            //var xx = iTextDocument.TextBuffer;

            var view = iWpfTextViewHost.TextView;
            var indexOfPublicInterface = view.TextSnapshot.GetText().IndexOf(interfaceName);
            if (indexOfPublicInterface > 0)
            {
                openFileWindow2.Activate();

                WriteLog("interface modification: start");

                view.Caret.MoveTo(view.GetTextViewLineContainingBufferPosition(view.TextSnapshot.GetLineFromPosition(0).Start));
                view.Caret.MoveTo(view.GetTextViewLineContainingBufferPosition(view.TextSnapshot.GetLineFromPosition(indexOfPublicInterface).End));
                System.Windows.Forms.SendKeys.SendWait("{DOWN}");
                System.Windows.Forms.SendKeys.SendWait("{LEFT}");

                System.Windows.Forms.SendKeys.SendWait(" : IAdoEntity");
                System.Threading.Thread.Sleep(200);
                System.Windows.Forms.SendKeys.SendWait("^S");
                System.Threading.Thread.Sleep(200);

                WriteLog("interface modification: end");
            }
        }

        private static object _lock = new object();

        private static void SimulateKeysToExtractInterface()
        {
            WriteLog("SimulateKeysToExtractInterface: start");

            lock (_lock)
            {

                // pass shortcut Ctrl + R; Ctrl + I to it.
                System.Windows.Forms.SendKeys.SendWait("^R^I");
                //System.Threading.Thread.Sleep(200);
                System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                //System.Threading.Thread.Sleep(100);
                System.Windows.Forms.SendKeys.SendWait("^S");
                System.Threading.Thread.Sleep(100);
                System.Windows.Forms.SendKeys.Flush();

                WriteLog("SimulateKeysToExtractInterface: end");
            }

            //keybd_event(Keys.ControlKey, 0, 0, 0);
            //keybd_event(Keys.ShiftKey, 0, 0, 0);
            //keybd_event(Keys.R, 0, 0, 0);
            //keybd_event(Keys.R, 0, KEYEVENTF_KEYUP, 0);
            //keybd_event(Keys.I, 0, 0, 0);

            //keybd_event(Keys.I, 0, KEYEVENTF_KEYUP, 0);
            //keybd_event(Keys.ShiftKey, 0, KEYEVENTF_KEYUP, 0);
            //keybd_event(Keys.ControlKey, 0, KEYEVENTF_KEYUP, 0);
            //System.Threading.Thread.Sleep(2000);

            //keybd_event(Keys.Enter, 0, 0, 0);
            //keybd_event(Keys.Enter, 0, KEYEVENTF_KEYUP, 0);
            //System.Threading.Thread.Sleep(2000);
        }

        private IWpfTextViewHost GetCurrentViewHost()

        {
            // code to get access to the editor's currently selected text cribbed from
            // http://msdn.microsoft.com/en-us/library/dd884850.aspx
            IVsTextManager txtMgr = (IVsTextManager)Package.GetGlobalService(typeof(SVsTextManager));
            int mustHaveFocus = 1;
            txtMgr.GetActiveView(mustHaveFocus, null, out IVsTextView vTextView);
            IVsUserData userData = vTextView as IVsUserData;
            if (userData == null)
            {
                return null;
            }
            else
            {
                Guid guidViewHost = Microsoft.VisualStudio.Editor.DefGuidList.guidIWpfTextViewHost;
                userData.GetData(ref guidViewHost, out object holder);
                IWpfTextViewHost viewHost = (IWpfTextViewHost)holder;
                return viewHost;
            }
        }

        public List<string> FindSubordinaryFiles(string sSourcePath, string[] excludeFolder)
        {
            var fileList = new List<string>();

            //Loop subordinary files
            DirectoryInfo theFolder = new DirectoryInfo(sSourcePath);
            FileInfo[] thefileInfo = theFolder.GetFiles("*.*", SearchOption.TopDirectoryOnly);

            foreach (FileInfo NextFile in thefileInfo)
                fileList.Add(NextFile.FullName);

            //Loop subordinary folders
            DirectoryInfo[] dirInfo = theFolder.GetDirectories();

            foreach (DirectoryInfo NextFolder in dirInfo.Where(di => !excludeFolder.Contains(di.Name)))
            {
                FileInfo[] fileInfo = NextFolder.GetFiles("*.*", SearchOption.AllDirectories);

                foreach (FileInfo NextFile in fileInfo)
                    fileList.Add(NextFile.FullName);
            }

            return fileList;
        }

        // Get Selected Text
        // https://stackoverflow.com/questions/2868127/get-the-selected-text-of-the-editor-window-visual-studio-extension

        private IWpfTextView GetWPFText()
        {
            var componentModel = (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
            var textManager = (IVsTextManager)Package.GetGlobalService(typeof(SVsTextManager));
            IVsTextView activeView = null;
            ErrorHandler.ThrowOnFailure(textManager.GetActiveView(1, null, out activeView));
            var editorAdapter = componentModel.GetService<IVsEditorAdaptersFactoryService>();
            IWpfTextView wpfTextView = editorAdapter.GetWpfTextView(activeView);

            return wpfTextView;
        }

        /// Given an IWpfTextViewHost representing the currently selected editor pane,
        /// return the ITextDocument for that view. That's useful for learning things 
        /// like the filename of the document, its creation date, and so on.
        private ITextDocument GetTextDocumentForView(IWpfTextViewHost viewHost)
        {
            ITextDocument document;
            viewHost.TextView.TextDataModel.DocumentBuffer.Properties.TryGetProperty(typeof(ITextDocument), out document);
            return document;
        }

        /// Get the current editor selection
        ITextSelection GetSelection(IWpfTextViewHost viewHost)
        {
            return viewHost.TextView.Selection;
        }

        //private static void LoopOpenDocumentLines(Document doc)
        //{
        //    ThreadHelper.ThrowIfNotOnUIThread();

        //    IVsEditorAdaptersFactoryService AdaptersFactory = null;

        //    Debug.Write("Activated : " + doc.FullName);
        //    var IvsTextView = GetIVsTextView(doc.FullName); //Calling the helper method to retrieve IVsTextView object.
        //    if (IvsTextView != null)
        //    {
        //        IvsTextView.GetBuffer(out curDocTextLines); //Getting Current Text Lines 

        //        //Getting Buffer Adapter to get ITextBuffer which holds the current Snapshots as wel..
        //        Microsoft.VisualStudio.Text.ITextBuffer curDocTextBuffer = AdaptersFactory.GetDocumentBuffer(curDocTextLines as IVsTextBuffer);
        //        Debug.Write("\r\nContentType: " + curDocTextBuffer.ContentType.TypeName + "\nTest: " + curDocTextBuffer.CurrentSnapshot.GetText());
        //    }
        //}

        public static void WriteLog(string msg)
        {
            string filePath = @"C:\Temp\Logs\VSExtension";
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }
            string logPath = @"C:\Temp\Logs\VSExtension\" + DateTime.Now.ToString("yyyy-MM-dd-hh") + ".txt";
            try
            {
                using (StreamWriter sw = File.AppendText(logPath))
                {
                    sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + msg);
                    sw.Flush();
                    sw.Close();
                    sw.Dispose();
                }
            }
            catch (IOException e)
            {
                using (StreamWriter sw = File.AppendText(logPath))
                {
                    sw.WriteLine("Exception：" + e.Message);
                    sw.WriteLine(DateTime.Now.ToString("yyy-MM-dd HH:mm:ss"));
                    sw.WriteLine("**************************************************");
                    sw.WriteLine();
                    sw.Flush();
                    sw.Close();
                    sw.Dispose();
                }
            }
        }
    }
}
