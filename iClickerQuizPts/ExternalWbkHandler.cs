using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Configuration;
using System.Data;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using iClickerQuizPts.AppExceptions;

namespace iClickerQuizPts
{
    /// <summary>
    /// Provides a mechanism for interacting with the iClicker-generated <see cref="Excel.Workbook"/> 
    /// containing the raw quiz data.
    /// </summary>
    public class ExternalWbkHandler
    {
        #region fields
        private readonly byte _firstDateCol; // ...to be read from App.config file
        private Excel.Workbook _wbkTestData = null;
        #endregion

        #region Ctor
        /// <summary>
        /// Instantiates a new instance of an <see cref="iClickerQuizPts.ExternalWbkHandler"/>.
        /// </summary>
        public ExternalWbkHandler()
        {
            //AppSettingsReader ar = new AppSettingsReader();
            //try
            //{
            //    _firstDateCol = (Byte)ar.GetValue("FirstDataCol", typeof(Byte));
            //}
            //catch(InvalidOperationException ex)
            //{

            //}
        }
        #endregion

        #region methods
        /// <summary>
        /// Prompts user to open XL wbk with latest iClick data.
        /// </summary>
        /// <returns>
        /// Returns name of opened XL workbook (string).  
        /// If user canceled out of FileDialog returns an empty string.
        /// </returns>
        public bool PromptUserToOpenQuizDataWbk()
        {
            bool userSelectedWbk = new bool();
            string testDataWbkNm = string.Empty;
            
            Office.FileDialog fd = Globals.ThisWorkbook.Application.get_FileDialog(
                Office.MsoFileDialogType.msoFileDialogOpen);
            fd.Title = "Latest iClick Results";
            fd.AllowMultiSelect = false;
            fd.Filters.Clear();
            fd.Filters.Add("Excel Files", "*.xlsx;*.xls");

            // Handle user selection...
            if (fd.Show() == -1) // ...-1 == file opened; 0 == user cxled
            {
                userSelectedWbk = true;
                fd.Execute();
                testDataWbkNm = Globals.ThisWorkbook.Application.ActiveWorkbook.Name;
                _wbkTestData = Globals.ThisWorkbook.Application.Workbooks[testDataWbkNm];
            }
            return userSelectedWbk;
        }

        /// <summary>
        /// Retrieves the contents of the column headers for those data columns which
        /// contain student quiz scores.
        /// </summary>
        /// <param name="noCols">The number of columns containing student quiz scores.
        /// <para>Do <i>not</i> include in this count columns for things like 
        /// student names or student email addresses.</para>
        /// </param>
        /// <returns>A string array of the contents of the quiz-data column headers.</returns>
        public string[] GetQuizFileHeaders(out long noCols)
        {
            Excel.Worksheet wsData = _wbkTestData.Worksheets[1];
            noCols =wsData.UsedRange.Columns.Count; // ...out param
            Excel.Range hdrs = wsData.UsedRange.Resize[1];
            string[] hdrContents = hdrs.Value2;
            return hdrContents;
        }

        /// <summary>
        /// Extracts the date portion of headers over the quiz-data-only columns
        /// of the iClicker-generated <see cref="Excel.Workbook"/> of raw 
        /// student quiz scores.
        /// </summary>
        /// <param name="headers">>A string array of the contents of the quiz-data column headers.</param>
        /// <param name="arrSize">The size of the array - i.e., the number of quiz-data column headers.</param>
        /// <returns>A generic <code>List&lt;T&gt;</code>, of type <code>DateTime</code>, 
        /// of the dates embedded in the headers of the quiz-data columns.</returns>
        public List<DateTime> GetQuizDatesFromHeaders(string[] headers, long arrSize)
        {
            List<DateTime> quizDates = new List<DateTime>();
            for(int i = _firstDateCol; i <= arrSize; i++)
            {
                quizDates.Add(GetDatePortionOfHeader(headers[i]));
            }
            return quizDates;
        }

        /// <summary>
        /// Extracts the date portion from the header cell of an iClicker data worksheet.
        /// </summary>
        /// <param name="hdr">The contents of a column header of an iClicker data worksheet.</param>
        /// <returns>The on which an iClicker quiz was given.</returns>
        /// <example>The text in the header cell is in a non-standard format and 
        /// therefore the enclosed date cannot be extracted.
        /// <para>NOTE: This exception will be thrown if this method is passed the text 
        /// from a cell which is not a header for quiz results (i.e., if the code is 
        /// pointing to an incorrect cell).</para></example>
        public DateTime GetDatePortionOfHeader(string hdr)
        {
            DateTime quizDate;
            try
            {
                hdr = hdr.Remove(1, "Session".Length + 1);
                hdr = hdr.Replace("Total ", string.Empty);
                hdr = hdr.Trim();
                // Hdr will now be something like:  "40 5/2/16 [2.00]"...
                int space1 = hdr.IndexOf(" ", 1);
                int space2 = hdr.IndexOf(" ", space1 +1);
                quizDate = DateTime.Parse(hdr.Substring(space1 + 1, space2 - space1 - 1));
                return quizDate;
            }
            catch(ParsingDateFmHdrException e)
            {
                e.HeaderText = hdr;
                throw e;
            }
        }
        #endregion
    }
}
