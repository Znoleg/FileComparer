using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using Microsoft.Office.Interop.Word;

namespace ComparerLevel1
{
    /// <summary>
    /// Allows you to get information about two files differences.
    /// Gets information only about modified lines.
    /// </summary>
    public class FileComparer
    {
        #region Fields

        private readonly string _originalFile;
        private readonly string _modifiedFile;
        private static readonly string[] txtFile = { ".txt" };
        private static readonly string[] docFile = { ".doc", ".docx"};
        private static readonly string[] pdfFile = { ".pdf" };
        private readonly Func<ICollection<string>> _action;

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

        #region Methods

        #region Public Methods

        /// <returns>Collection of modified lines in "index: modified line value for line index" format</returns>
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

            var originalText = original.ReadToEnd().Split('\n'); // read from text file
            var modifiedText = modified.ReadToEnd().Split('\n');

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
            // read from doc file
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
            // read from pdf file
            for (int i = 1; i <= originalPDF.GetNumberOfPages(); i++)
            {
                string originalPage = PdfTextExtractor.GetTextFromPage(originalPDF.GetPage(i));
                originalText.AddRange(originalPage.Split('\n'));
            }
            for (int i = 1; i <= modifiedPDF.GetNumberOfPages(); i++)
            {
                string modifiedPage = PdfTextExtractor.GetTextFromPage(modifiedPDF.GetPage(i));
                modifiedText.AddRange(modifiedPage.Split('\n'));
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
            ICollection<string> changes = new List<string>();
            int origLen = originalText.Count;
            int modifiedLen = modifiedText.Count;
            int minLen = Math.Min(origLen, modifiedLen);
            int maxLen = Math.Max(origLen, modifiedLen);

            for (int i = 0; i < minLen; i++) // Just consistently compares two strings - if string not equal they're modified
            {
                if (originalText[i] != modifiedText[i]) changes.Add($"{i + 1}: <modified line value for line {i + 1}>");
            }
            for (int i = minLen; i < maxLen; i++) // Are added lines considered as modified? if not this functionality can be commented out
            {
                changes.Add($"{i + 1}: <modified line value for line {i + 1}>");
            }

            return changes;
        }

        /// <summary>
        /// Gets right funcion depends on file type arrays of class
        /// </summary>
        /// <param name="originalFile"></param>
        /// <param name="modifiedFile"></param>
        private Func<ICollection<string>> GetAction(string originalFile, string modifiedFile)
        {
            string originalType = GetFileType(originalFile);
            string modifiedType = GetFileType(modifiedFile);
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

        #endregion Private Methods

        #endregion Methods
    }
}
