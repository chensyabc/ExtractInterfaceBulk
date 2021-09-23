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
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;
using SystemConfiguration = System.Configuration;
using static ExtractInterfaceBulk.VSExtensionHelper;

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


        #region Members

        private string SolutionDirectory { get; set; }
        private string MoveInterfaceToFolder = "Shared/Interfaces";
        private string ProjectName { get; set; }
        private string[] RemoveLines { get; set; }
        private string[] DeleteLinesWildCard { get; set; }
        private string[] SpecialClassesWithBaseInterface { get; set; }
        private string[] IgnoreClasses { get; set; }

        #endregion


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

            var ignoreFolders = new string[] { "bin", "obj", "Interfaces" };
            var baseInterface = "IBaseInterface";

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
            var projectFileName = Path.GetFileName(project.FullName);
            var projectName = Path.GetFileNameWithoutExtension(project.FullName);
            var projectDir = Path.GetDirectoryName(project.FullName);

            var needAddClasses = new List<string>();
            var alreadyHasInterfaceClasses = new List<string>();
            var allFiles = FindSubordinaryFiles(projectDir, ignoreFolders);

            this.SolutionDirectory = solutionDir;
            this.ProjectName = projectName;

            string[] lines = File.ReadAllLines(@"Setting/DeleteLines.txt");
            var deleteLines = lines.Where(line => !string.IsNullOrEmpty(line)).ToArray();

            string[] linesWildCard = File.ReadAllLines(@"Setting/DeleteLinesWildCard.txt");
            this.DeleteLinesWildCard = linesWildCard.Where(line => !string.IsNullOrEmpty(line)).ToArray();

            string[] specialClassesWithBaseInterface = File.ReadAllLines(@"Setting/SpecialClassesWithBaseInterface.txt");
            this.SpecialClassesWithBaseInterface = specialClassesWithBaseInterface.Where(line => !string.IsNullOrEmpty(line)).ToArray();

            string[] ignoreClasses = File.ReadAllLines(@"Setting/IgnoreClasses.txt");
            this.IgnoreClasses = ignoreClasses.Where(line => !string.IsNullOrEmpty(line)).ToArray();

            // OptionA: extract interface bulk
            allFiles.ForEach(file =>
            {
                var processedClass = ExtractInterface(file, vsDte2, alreadyHasInterfaceClasses, baseInterface, deleteLines);

                if (!string.IsNullOrEmpty(processedClass))
                    needAddClasses.Add(processedClass);
            });

            WriteLog("Writing final summary message.");
            var summaryMsg = string.Format(@"Processed Classes: {0} \n\r Already has Interface: {1}", string.Join(",", needAddClasses.ToArray()), string.Join(",", alreadyHasInterfaceClasses.ToArray()));
            WriteLog(summaryMsg);
            WriteLog("Execution End.");

            // Show message box about Processed Classes
            VsShellUtilities.ShowMessageBox(
                this.package,
                summaryMsg,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        /// <summary>
        /// return processed class name
        /// </summary>
        private string ExtractInterface(string filePathWithExtension, DTE2 vsDte2, List<string> alreadyHasInterfaceClasses, string baseInterface, string[] deleteLines)
        {
            var processedClass = string.Empty;
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            // only process class files, and .cs is first judgement
            if (filePathWithExtension.EndsWith(".cs"))
            {
                // 1, Open File
                var openFileWindow = OpenFileAndActive(vsDte2, filePathWithExtension);

                // 2, Get Editing View
                var editingView = GetEditingView();
                var allText = editingView.TextSnapshot.GetText();
                //Keep below, this may be used in future.
                //var iTextDocument = GetTextDocumentForView(iWpfTextViewHost);
                //var xx = iTextDocument.TextBuffer;

                // 3, Move to Caret of interface line
                var textToSearch = "public class";
                if (!MoveToCaretToSearchText(editingView, textToSearch))
                {
                    return processedClass;
                }
                openFileWindow.Activate();

                // 4, Find class line
                Regex regex = new Regex(@"public class (\S+)");
                if (regex.IsMatch(allText))
                {
                    processedClass = regex.Match(allText).Value.Substring(13);
                }
                else
                {
                    WriteLog("find no class: " + processedClass);
                    return string.Empty;
                }

                WriteLog("find class: " + processedClass);
                if (this.IgnoreClasses.Contains(processedClass))
                    return string.Empty;

                // 5, Check if it has interface
                var interfaceName = string.Format("I{0}", processedClass);
                var interfacePart1 = string.Format(": I{0}", processedClass);
                var interfacePart2 = string.Format(", I{0}", processedClass);
                if (allText.Contains(interfacePart1) || allText.Contains(interfacePart2))
                {
                    WriteLog("already extract interface: " + processedClass);
                    alreadyHasInterfaceClasses.Add(processedClass);
                    return string.Empty;
                }

                // 6, extract interface for class
                openFileWindow.Activate();
                SimulateKeysToExtractInterface();

                var replaceBaseInterfaceResult = ReplaceContent(editingView.TextBuffer, null, $" {baseInterface},", string.Empty);

                var parentDirectory = Path.GetDirectoryName(filePathWithExtension);
                string interfaceFile = Path.Combine(parentDirectory, $"{interfaceName}.cs");
                if (!File.Exists(interfaceFile))
                {
                    WriteLog("interface file not found!");
                    return string.Empty;
                }

                // 7, modify interface
                ModifyInterface(interfaceFile, vsDte2, new List<string>(), baseInterface, replaceBaseInterfaceResult || this.SpecialClassesWithBaseInterface.Contains(processedClass), deleteLines);

                // 8, move interface file
                MoveInterface(interfaceFile, this.MoveInterfaceToFolder);

                openFileWindow.Close(vsSaveChanges.vsSaveChangesYes);

                WriteLog("extract interface successfully: " + processedClass);
            }

            return processedClass;
        }

        private void MoveInterface(string interfaceFile, Func<string, string, string> moveInterfaceToFolder)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// return processed interface name
        /// </summary>
        private string ModifyInterface(string filePathWithExtension, DTE2 vsDte2, List<string> alreadyHasInterfaceClasses, string baseInterface, bool isClassInheritBaseInterface, string[] deleteLines)
        {
            var processedInterface = string.Empty;
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            // only process interface files, and .cs is first judgement
            if (filePathWithExtension.EndsWith(".cs"))
            {
                WriteLog("begin ModifyInterface file: " + filePathWithExtension);

                // 1, Open File
                var openFileWindow = OpenFileAndActive(vsDte2, filePathWithExtension);

                // 2, Get Editing View
                var editingView = GetEditingView();
                var allText = editingView.TextSnapshot.GetText();

                // 3, Move to Caret of interface line
                if (!MoveToCaretToSearchText(editingView, "public interface"))
                {
                    return processedInterface;
                }
                openFileWindow.Activate();

                bool isToAddBaseInterface = true;
                bool isToReplaceNameSpace = true;
                bool isToRemoveLines = true;

                // add base interface
                if (isToAddBaseInterface && isClassInheritBaseInterface)
                {
                    // 4, See if Interface has Inherit interfaceInherit
                    var interfacePart1 = string.Format(": {0}", baseInterface);
                    var interfacePart2 = string.Format(", {0}", baseInterface);
                    if (allText.Contains(interfacePart1) || allText.Contains(interfacePart2))
                    {
                        WriteLog("already inherit interface: " + processedInterface);
                        alreadyHasInterfaceClasses.Add(processedInterface);
                        return string.Empty;
                    }

                    // 5, Get Interface Name
                    Regex regex = new Regex(@"public interface (\S+)");
                    if (regex.IsMatch(allText))
                    {
                        processedInterface = regex.Match(allText).Value.Substring(17);
                        WriteLog("find interface: " + processedInterface);
                    }
                    else
                    {
                        WriteLog("find no interface: " + processedInterface);
                        return processedInterface;
                    }

                    // Ignore BaseInterface
                    if (processedInterface == baseInterface)
                        return string.Empty;

                    // 6, Append Interface interfaceInherit
                    openFileWindow.Activate();
                    SimulateKeysToMakeInterfaceInheritBaseInterface(baseInterface);
                }

                // Replace namespace
                if (isToReplaceNameSpace)
                {
                    // 7.0, if previous already changes, modify
                    if (allText.Contains($"namespace {this.ProjectName}"))
                    {
                        ReplaceContent(editingView.TextBuffer, null, $"namespace {this.ProjectName}", "namespace Shared.Interfaces");
                    }
                    else
                    {
                        alreadyHasInterfaceClasses.Add(processedInterface);
                        return string.Empty;
                    }
                }

                // Remove lines
                if (isToRemoveLines)
                {
                    if (isClassInheritBaseInterface)
                    {
                        if (deleteLines.Any())
                        {
                            DeleteLine(editingView.TextBuffer, null, deleteLines);
                        }
                    }

                    foreach (var deleteLineWildCard in this.DeleteLinesWildCard)
                    {
                        DeleteLineWithWildCard(editingView.TextBuffer, null, deleteLineWildCard);
                    }
                }

                if (isClassInheritBaseInterface)
                {
                    InsertAfterSpecificText(editingView, "    {", "        #region Members\r\n\r\n");

                    var appendLines = "        #endregion\r\n";
                    appendLines += "\r\n\r\n";
                    appendLines += "        #region NonDB Members\r\n        #endregion\r\n";
                    appendLines += "\r\n\r\n";
                    appendLines += "        #region Navigation Properties\r\n        #endregion\r\n";
                    //appendLines += "\r\n\r\n";
                    //appendLines += "        #region Manipulations\r\n        #endregion\r\n";
                    AppendPreviousLineInFile(editingView, "    }", appendLines);
                }
                else
                {
                    InsertAfterSpecificText(editingView, "    {", "        #region Members\r\n\r\n");

                    var appendLines = "        #endregion\r\n";
                    appendLines += "\r\n\r\n";
                    appendLines += "        #region Navigation Properties\r\n        #endregion\r\n";
                    AppendPreviousLineInFile(editingView, "    }", appendLines);
                }

                openFileWindow.Close(vsSaveChanges.vsSaveChangesYes);

                WriteLog("end ModifyInterface successfully: " + processedInterface);
            }

            return processedInterface;
        }

        private string MoveInterface(string filePathWithExtension, string moveInterfaceToFolder)
        {
            if (string.IsNullOrEmpty(moveInterfaceToFolder)) return string.Empty;

            var processedInterface = string.Empty;
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            // only process interface files, and .cs is first judgement
            if (filePathWithExtension.EndsWith(".cs"))
            {
                WriteLog("begin ModifyInterface file: " + filePathWithExtension);

                // 1, see if file exists
                if (File.Exists(filePathWithExtension))
                {
                    try
                    {
                        var fileNameWithExtension = Path.GetFileName(filePathWithExtension);
                        var directoryName = new DirectoryInfo(Path.GetDirectoryName(filePathWithExtension)).Name;
                        var newPath = Path.Combine(this.SolutionDirectory, moveInterfaceToFolder);
                        if (directoryName != this.ProjectName)
                        {
                            var parentParentDirectoryName = new DirectoryInfo(Directory.GetParent(Path.GetDirectoryName(filePathWithExtension)).FullName).Name;

                            if (parentParentDirectoryName != this.ProjectName)
                            {
                                newPath = Path.Combine(this.SolutionDirectory, moveInterfaceToFolder, parentParentDirectoryName, directoryName);
                            }
                            else
                            {
                                newPath = Path.Combine(this.SolutionDirectory, moveInterfaceToFolder, directoryName);
                            }
                        }
                        var newFilePathWithExtension = Path.Combine(newPath, fileNameWithExtension);

                        if (!Directory.Exists(newPath)) Directory.CreateDirectory(newPath);

                        if (!File.Exists(newFilePathWithExtension))
                        {
                            File.Move(filePathWithExtension, newFilePathWithExtension);
                        }
                    }
                    catch (Exception e)
                    {
                        throw;
                    }
                }

                WriteLog("end ModifyInterface successfully: " + processedInterface);
            }

            return processedInterface;
        }


        private bool MoveToCaretToSearchText(IWpfTextView view, string textToSearch)
        {
            var indexOfPublicClass = view.TextSnapshot.GetText().IndexOf(textToSearch);
            if (indexOfPublicClass < 0)
                return false;
            view.Caret.MoveTo(view.GetTextViewLineContainingBufferPosition(view.TextSnapshot.GetLineFromPosition(0).Start));
            view.Caret.MoveTo(view.GetTextViewLineContainingBufferPosition(view.TextSnapshot.GetLineFromPosition(indexOfPublicClass).Start));

            return true;
        }

        private Window OpenFileAndActive(DTE2 vsDte2, string filePathWithExtension)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // Open interface file
            WriteLog("before open file: " + filePathWithExtension);
            var openFileWindow = vsDte2.ItemOperations.OpenFile(filePathWithExtension);
            openFileWindow.Activate();
            System.Threading.Thread.Sleep(100);
            WriteLog("opened file: " + filePathWithExtension);

            return openFileWindow;
        }

        private static object _lock = new object();

        private static void SimulateKeysToExtractInterface()
        {
            lock (_lock)
            {
                // pass shortcut Ctrl + R; Ctrl + I to it.
                System.Windows.Forms.SendKeys.Send("^R^I");
                //System.Threading.Thread.Sleep(200);
                System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                //System.Threading.Thread.Sleep(100);
                System.Windows.Forms.SendKeys.SendWait("^S");
                System.Threading.Thread.Sleep(1);
                System.Windows.Forms.SendKeys.Flush();
            }
        }

        private void SimulateKeysToMakeInterfaceInheritBaseInterface(string baseInterface)
        {
            lock (_lock)
            {
                // pass shortcut Ctrl + R; Ctrl + I to it.
                SendKeys.SendWait("{END}");
                SendKeys.SendWait(string.Format(" : {0}", baseInterface));
                //SendKeys.SendWait("^S");
                SendKeys.Flush();
                System.Threading.Thread.Sleep(1);
            }
        }

        private IWpfTextView GetEditingView()
        {
            var iWpfTextViewHost = GetCurrentViewHost();
            var editingView = iWpfTextViewHost.TextView;

            return editingView;
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
