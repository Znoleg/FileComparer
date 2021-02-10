using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using Microsoft.Office.Interop.Word;
using NLog;

namespace ComparerLevel3
{
    /// <summary>
    /// Allows you to get information about two files differences.
    /// Gets information about deleted and added lines.
    /// </summary>
    public class FileComparer
    {
        #region Fields

        private readonly string _originalFile;
        private readonly string _modifiedFile;
        private static readonly string[] txtFile = { ".txt" };
        private static readonly string[] docFile = { ".doc", ".docx" };
        private static readonly string[] pdfFile = { ".pdf" };
        private readonly Func<ICollection<string>> _action;
        private readonly Logger _logger = null;

        #endregion

        /// <param name="originalFile">Original file path</param>
        /// <param name="modifiedFile">Modified file path</param>
        public FileComparer(string originalFile, string modifiedFile)
        {
            if (originalFile == modifiedFile) throw new ArgumentException("File paths are the same");
            if (!File.Exists(originalFile)) throw new FileNotFoundException($"Following file not found: { originalFile }");
            if (!File.Exists(originalFile)) throw new FileNotFoundException($"Following file not found: { modifiedFile }");
            _action = GetAction(originalFile, modifiedFile);

            _originalFile = originalFile;
            _modifiedFile = modifiedFile;
        }

        /// <param name="originalFile">Original file path</param>
        /// <param name="modifiedFile">Modified file path</param>
        /// <param name="logger">NLog logger that will log string value and its status</param>
        public FileComparer(string originalFile, string modifiedFile, Logger logger) : this(originalFile, modifiedFile)
        {
            _logger = logger;
            _logger.Info($"{originalFile} and {modifiedFile} files: \n");
        }

        #region Methods

        #region Public Methods

        /// <returns>Collection of modified lines in "index: added line /+ removed line" format</returns>
        public ICollection<string> GetDifference()
        {
            return _action.Invoke();
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Gets texts from Txt file and calls returns their differences
        /// </summary>
        private ICollection<string> CompareTxt()
        {
            StreamReader original = new StreamReader(_originalFile);
            StreamReader modified = new StreamReader(_modifiedFile);

            var originalText = original.ReadToEnd().Split("\n");
            var modifiedText = modified.ReadToEnd().Split("\n");

            original.Close();
            modified.Close();

            return CompareTexts(originalText, modifiedText);
        }

        /// <summary>
        /// Gets texts from Word file and calls returns their differences
        /// </summary>
        private ICollection<string> CompareWord()
        {
            Application word = new Application();
            object miss = System.Reflection.Missing.Value;
            object readOnly = true;
            Document original = word.Documents.Open(_originalFile, ref miss, ref readOnly);
            Document modified = word.Documents.Open(_modifiedFile, ref miss, ref readOnly);

            List<string> originalText = new List<string>();
            List<string> modifiedText = new List<string>();
            for (int i = 0; i < original.Paragraphs.Count; i++) originalText.Add(original.Paragraphs[i + 1].Range.Text.ToString());
            for (int i = 0; i < modified.Paragraphs.Count; i++) modifiedText.Add(modified.Paragraphs[i + 1].Range.Text.ToString());

            original.Close();
            modified.Close();

            return CompareTexts(originalText, modifiedText);
        }

        /// <summary>
        /// Gets texts from PDF file and calls returns their differences
        /// </summary>
        private ICollection<string> ComparePDF()
        {
            var originalPDF = new PdfDocument(new PdfReader(_originalFile));
            var modifiedPDF = new PdfDocument(new PdfReader(_modifiedFile));

            List<string> originalText = new List<string>();
            List<string> modifiedText = new List<string>();
            for (int i = 1; i <= originalPDF.GetNumberOfPages(); i++)
            {
                string originalPage = PdfTextExtractor.GetTextFromPage(originalPDF.GetPage(i));
                originalText.AddRange(originalPage.Split("\n"));
            }
            for (int i = 1; i <= modifiedPDF.GetNumberOfPages(); i++)
            {
                string modifiedPage = PdfTextExtractor.GetTextFromPage(modifiedPDF.GetPage(i));
                modifiedText.AddRange(modifiedPage.Split("\n"));
            }

            originalPDF.Close();
            modifiedPDF.Close();

            return CompareTexts(originalText, modifiedText);
        }

        /// <summary>
        /// Finds differences between "originalText" and "modifiedText"
        /// </summary>
        /// <param name="originalText">String collection (text) of original file</param>
        /// <param name="modifiedText">String collection (text) of modified file</param>
        private ICollection<string> CompareTexts(IReadOnlyList<string> originalText, IReadOnlyList<string> modifiedText)
        {
            List<string> changes = new List<string>();
            const string addedStatus = " <added line to modified>";
            const string removedStatus = " <removed line from original>";

            int maxLen = originalText.Count >= modifiedText.Count ? originalText.Count : modifiedText.Count;
            for (int i = 0; i < maxLen; i++) changes.Add(""); // Adding all strings to the collection - to easily add status lines
            // (text line index match collection index)

            Queue<int> viewed = new Queue<int>();
            // There are two different indexes for origFile string and for modifiedFile - they may disperse later
            int modIndx = 0;
            for (int origIndx = 0; origIndx < originalText.Count; origIndx++, modIndx++)
            {
                bool stringFound = false;
                if (originalText[origIndx] != modifiedText[modIndx]) // Just consistently compares two strings
                {
                    // if not equal 
                    // we must find the line of the original file in the modified file - it could just move lower
                    for (int j = modIndx; j < modifiedText.Count; j++)
                    {
                        if (originalText[origIndx] != modifiedText[j]) // until we find it, add the viewed lines to the queue
                        {
                            viewed.Enqueue(j);
                        }
                        else // if we found it (line just lowered because of new added lines)
                        {
                            while (viewed.Count > 0) // adding "added" status to this new lines from queue
                            {
                                int index = viewed.Dequeue();
                                AppendStringStatus(changes, index, addedStatus, modifiedText[index]);
                                modIndx = j; // we looked at these lines and realized that they are added - we can skip them
                            }
                            stringFound = true; // flag
                            break;
                        }
                    }

                    if (!stringFound) // string not found - string is removed
                    {
                        viewed.Clear();
                        AppendStringStatus(changes, origIndx, removedStatus, originalText[origIndx]);
                        modIndx--; // to set this string "added" status next - not only removed
                    }
                }
            }

            for (int i = modIndx; i < modifiedText.Count; i++) // additional string is added
            {
                AppendStringStatus(changes, i, addedStatus, modifiedText[i]);
            }

            for (int i = 0, j = 0; i < changes.Count; j++) // clearing out empty strings and set right string index like "3: <added line>"
            {
                if (changes[i] != "")
                {
                    changes[i] = changes[i].Insert(0, (j + 1) + ":");
                    i++;
                }

                else changes.RemoveAt(i);
            }

            return changes;
        }

        /// <summary>
        /// Gets right funcion depends on file type arrays of class
        /// </summary>
        /// <param name="originalFile"></param>
        /// <param name="modifiedFile"></param>
        private Func<ICollection<string>> GetAction(string file1, string file2)
        {
            string originalType = GetFileType(file1);
            string modifiedType = GetFileType(file2);
            if (originalType != modifiedType) throw new ArgumentException("Files type aren't equal");

            if (txtFile.Any(str => str == originalType)) return CompareTxt;
            if (docFile.Any(str => str == originalType)) return CompareWord;
            if (pdfFile.Any(str => str == originalType)) return ComparePDF;

            else throw new ArgumentException("Unknown file format");
        }

        /// <summary>
        /// Cuts provided string to the last '.'
        /// </summary>
        /// <returns>String in ".type" format</returns>
        private string GetFileType(string file)
        {
            int typeIndex = file.LastIndexOf('.');
            return file.Substring(typeIndex);
        }

        /// <summary>
        /// Appends "status" string to "toAppendIn" collection and logs info to logger if exists
        /// </summary>
        /// <param name="inText">String value from file</param>
        private void AppendStringStatus(IList<string> toAppendIn, int index, string status, string inText)
        {
            toAppendIn[index] += status;
            if (_logger != null) Log(status, inText);
        }

        private void Log(string status, string inText)
        {
            if (status.Contains("\r")) status = status.Replace("\r", "");
            if (inText.Contains("\r")) inText = inText.Replace("\r", "");
            _logger.Info($"{inText}: {status}");
        }

        #endregion Private Methods

        #endregion Methods
    }
}
