﻿using System;
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
    public class ExternalWbkHandler
    {
        #region fields
        private readonly byte _firstDateCol; // ...to be read from App.config file
        private Excel.Workbook _wbkTestData = null;
        #endregion

        #region Ctor
        ExternalWbkHandler()
        {
            AppSettingsReader ar = new AppSettingsReader();
            _firstDateCol = (Byte)ar.GetValue("FirstDataCol",typeof(Byte));
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

        public string[] GetQuizFileHeaders(out long noCols)
        {
            Excel.Worksheet wsData = _wbkTestData.Worksheets[1];
            noCols =wsData.UsedRange.Columns.Count; // ...out param
            Excel.Range hdrs = wsData.UsedRange.Resize[1];
            string[] hdrContents = hdrs.Value2;
            return hdrContents;
        }

        public List<DateTime> GetQuizDatesFromHeaders(string[] headers, long arrSize)
        {
            List<DateTime> quizDates = new List<DateTime>();
            for(int i = _firstDateCol; i <= arrSize; i++)
            {
                quizDates.Add(GetDatePortionOfHeader(headers[i]));
            }
            return quizDates;
        }

        public DateTime GetDatePortionOfHeader(string hdr)
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
        #endregion
    }
}