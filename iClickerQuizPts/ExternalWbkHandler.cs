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
using OfficeOpenXml;
using System.IO;

namespace iClickerQuizPts
{
    public static class ExternalWbkHandler
    {
        #region fields
        static DataTable dtQz = null;
        // Values to be read from App.config file...
        public readonly static byte ColNoDataBegXL;
        public readonly static byte ColNoEmlXL;
        public readonly static byte ColNoStdntNmXL;
        public readonly static string ColNmID;
        public readonly static string ColNmEmail;
        public readonly static string ColNmLNm;
        public readonly static string ColNmFNm;
        public readonly static byte NmbrNonScoreColsDT;
        #endregion

        #region Ppts
        public static Excel.Workbook WbkTestData { get; set; } = null;
        public static DataTable QuizDataTable
        {
            get { return dtQz; }
        }
        #endregion

        #region Ctor
        static ExternalWbkHandler()
        {
            // Read values from app.config file into public fields...
            AppSettingsReader ar = new AppSettingsReader();
            ColNoEmlXL = (byte)ar.GetValue("ColNoEmailXL", typeof(byte));
            ColNoStdntNmXL = (byte)ar.GetValue("ColNoStdntNmXL", typeof(byte));
            ColNoDataBegXL = (byte)ar.GetValue("ColNoDataBegins",typeof(Byte));
            ColNmID = (string)ar.GetValue("Col00", typeof(string));
            ColNmEmail = (string)ar.GetValue("Col01", typeof(string));
            ColNmLNm = (string)ar.GetValue("Col02", typeof(string));
            ColNmFNm = (string)ar.GetValue("Col03", typeof(string));
            NmbrNonScoreColsDT = (byte)ar.GetValue("NmbrNonScoreCols", typeof(byte));
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
        public static bool PromptUserToOpenQuizDataWbk()
        {
            bool userSelectedWbk = new bool();
            string testDataWbkNm = string.Empty;
            
            Office.FileDialog fd = Globals.ThisWorkbook.Application.get_FileDialog(
                Office.MsoFileDialogType.msoFileDialogFilePicker);
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
                WbkTestData = Globals.ThisWorkbook.Application.Workbooks[testDataWbkNm];
            }
            return userSelectedWbk;
        }

        public static string[] GetQuizFileHeaders(out long noTtlHdrs)
        {
            Excel.Worksheet wsData = WbkTestData.Worksheets[1];
            noTtlHdrs =wsData.UsedRange.Columns.Count; // ...out param
            Excel.Range hdrs = wsData.UsedRange.Resize[1];
            string[] hdrContents = hdrs.Value2;
            return hdrContents;
        }

        public static List<DateTime> GetQuizDatesFromHeaders(string[] headers, long arrSize)
        {
            List<DateTime> quizDates = new List<DateTime>();
            for(int i = ColNoDataBegXL; i <= arrSize; i++)
            {
                quizDates.Add(GetDatePortionOfHeader(headers[i]));
            }
            return quizDates;
        }

        public static DateTime GetDatePortionOfHeader(string hdr)
        {
            DateTime quizDate;
            try
            {
                hdr.Remove(1, "Session".Length + 1);
                hdr.Replace("Total ", string.Empty);
                hdr.Trim();
                // Hdr will now be something like:  "40 5/2/16 [2.00]"...
                int space1 = hdr.IndexOf(" ", 1);
                int space2 = hdr.IndexOf(" ", space1 + 1);
                quizDate = DateTime.Parse(hdr.Substring(space1 + 1, space2 - space1 - 1));
                return quizDate;
            }
            catch(ParsingDateFmHdrException e)
            {
                e.HeaderText = hdr;
                throw e;
            }
        }

        public static void AddEmptyNonQuizColumnsToDataTable()
        {
            // Make sure data table is 'empty'...
            if (QuizDataTable.Columns.Count != 0)
            {
                QuizDataTable.Clear();
                QuizDataTable.Columns.Clear();
            }

            DataColumn dcID = new DataColumn(ColNmID, typeof(int));
            dcID.AutoIncrement = true;
            dcID.ReadOnly = true;
            QuizDataTable.Columns.Add(dcID);

            DataColumn dcEml = new DataColumn(ColNmEmail, typeof(string));
            dcEml.AllowDBNull = false;
            dcEml.Unique = true;
            dcEml.MaxLength = 60;
            QuizDataTable.Columns.Add(dcEml);

            DataColumn dcLN = new DataColumn(ColNmLNm, typeof(string));
            dcLN.AllowDBNull = true;
            dcLN.Unique = false;
            dcLN.MaxLength = 30;
            QuizDataTable.Columns.Add(dcLN);

            DataColumn dcFN = new DataColumn(ColNmFNm, typeof(string));
            dcFN.AllowDBNull = true;
            dcFN.Unique = false;
            dcFN.MaxLength = 30;
            QuizDataTable.Columns.Add(ColNmFNm, typeof(string));
        }

        public static void ReadXLFileIntoDataTable(string wbNm)
        {
            ExcelPackage package = new ExcelPackage(new FileInfo(wbNm));
            ExcelWorksheet ws = package.Workbook.Worksheets[1];
            ExcelRow rowHdrs = ws.Row(1);
            int nmbrScoreColzXL = 1 + ws.Dimension.End.Column - ColNoDataBegXL;

            // Add columns & their headers for quiz data...
            for(int i = ColNoDataBegXL; i <=ws.Dimension.End.Column; i++)
            {
                string colNm = ws.Cells[1, i].Text;
                colNm = GetDatePortionOfHeader(colNm).ToShortDateString();
                DataColumn dc = new DataColumn(colNm, typeof(byte));
                dc.AllowDBNull = true;
                dc.Unique = false;
                QuizDataTable.Columns.Add(dc);
            }

            // Now add the data into the DataTable...
            for (int i = ws.Dimension.Start.Row + 1; i <= ws.Dimension.End.Row; i++)
            {
                DataRow dr = QuizDataTable.NewRow();
                string stNm = ws.Cells[i, ColNoStdntNmXL].Text.Trim();
                int commaLoc = stNm.IndexOf(","); // ... -1 if not present

                // Populate names & email cols...
                dr[ColNmEmail] = ws.Cells[i, ColNoEmlXL].Text;
                switch (commaLoc)
                {
                    case -1:
                        dr[ColNmLNm] = stNm;
                        break;
                    case 0:
                        dr[ColNmLNm] = String.Empty;
                        dr[ColNmFNm] = stNm.Substring(1);
                        break;
                    default:
                        dr[ColNmLNm] = stNm.Substring(0, commaLoc);
                        dr[ColNmFNm] = stNm.Substring(commaLoc + 1);
                        break;
                }

                // Populate data cols...
                for(int j = 1; j <= nmbrScoreColzXL; j++)
                {
                    dr[NmbrNonScoreColsDT + j] = float.Parse(ws.Cells[i, ColNoDataBegXL + j -1].Text);
                }

                // Add row...
                QuizDataTable.Rows.Add(dr);
            }
        }
        #endregion
    }
}
