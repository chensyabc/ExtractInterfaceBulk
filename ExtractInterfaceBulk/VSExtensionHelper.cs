using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using System.Linq;

namespace ExtractInterfaceBulk
{
    public static class VSExtensionHelper
    {
        public static void AppendPreviousLineInFile(IWpfTextView editingView, string toBeAppend, string appendment)
        {
            var textBuffer = editingView.TextBuffer;
            var textEdit = textBuffer.CreateEdit();
            if (textBuffer.CheckEditAccess())
            {
                var snapshot = textEdit.Snapshot;
                for (int i = 0; i < snapshot.Lines.Count(); i++)
                {
                    var line = snapshot.Lines.ElementAt(i);
                    if (line.LineBreakLength != 0)
                    {
                        string breakText = line.GetLineBreakText();
                        string pureText = line.GetText();
                        if (pureText == toBeAppend)
                        {
                            textEdit.Insert(line.Start, appendment);
                        }
                    }
                }

                textEdit.Apply();
            }
        }

        public static void InsertAfterSpecificText(IWpfTextView editingView, string toBeInserted, string increasement)
        {
            var textBuffer = editingView.TextBuffer;
            var textEdit = textBuffer.CreateEdit();
            if (textBuffer.CheckEditAccess())
            {
                var snapshot = textEdit.Snapshot;
                foreach (ITextSnapshotLine line in snapshot.Lines)
                {
                    if (line.LineBreakLength != 0)
                    {
                        string breakText = line.GetLineBreakText();
                        string pureText = line.GetText();
                        if (pureText == toBeInserted)
                        {
                            //textEdit.Replace(line.End + line.LineBreakLength, 0, replacement);
                            textEdit.Insert(line.End + line.LineBreakLength, increasement);
                        }
                    }
                }

                textEdit.Apply();
            }
        }

        public static bool ReplaceContent(ITextBuffer textBuffer, ITextEdit textEdit, string toBeReplaced, string replacement)
        {
            bool isSuccess = false;
            textEdit = textBuffer.CreateEdit();

            if (textBuffer.CheckEditAccess())
            {
                var snapshot = textEdit.Snapshot;
                foreach (ITextSnapshotLine line in snapshot.Lines)
                {
                    if (line.LineBreakLength != 0)
                    {
                        string pureText = line.GetText();
                        if (pureText.Contains(toBeReplaced))
                        {
                            var index = pureText.IndexOf(toBeReplaced);
                            var replaceResult = textEdit.Replace(line.Start + index, toBeReplaced.Length, replacement);
                            isSuccess = replaceResult;
                        }
                    }
                }

                textEdit.Apply();
            }

            return isSuccess;
        }

        public static bool DeleteLine(ITextBuffer textBuffer, ITextEdit textEdit, string[] deleteLines)
        {
            bool isSuccess = false;
            textEdit = textBuffer.CreateEdit();

            if (textBuffer.CheckEditAccess())
            {
                var snapshot = textEdit.Snapshot;
                foreach (ITextSnapshotLine line in snapshot.Lines)
                {
                    if (line.LineBreakLength != 0)
                    {
                        string pureText = line.GetText();
                        string pureTextTrim = pureText.Trim();
                        if (deleteLines.Contains(pureTextTrim))
                        {
                            var deleteResult = textEdit.Delete(line.Start, line.LengthIncludingLineBreak);
                            isSuccess = deleteResult;
                        }
                    }
                }

                textEdit.Apply();
            }

            return isSuccess;
        }

        public static bool DeleteLineWithWildCard(ITextBuffer textBuffer, ITextEdit textEdit, string wildCardStartWith)
        {
            bool isSuccess = false;
            textEdit = textBuffer.CreateEdit();

            if (textBuffer.CheckEditAccess())
            {
                var snapshot = textEdit.Snapshot;
                foreach (ITextSnapshotLine line in snapshot.Lines)
                {
                    if (line.LineBreakLength != 0)
                    {
                        string pureText = line.GetText();
                        string pureTextTrim = pureText.Trim();
                        if (pureTextTrim.StartsWith(wildCardStartWith))
                        {
                            var deleteResult = textEdit.Delete(line.Start, line.LengthIncludingLineBreak);
                            isSuccess = deleteResult;
                        }
                    }
                }

                textEdit.Apply();
            }

            return isSuccess;
        }
    }
}
