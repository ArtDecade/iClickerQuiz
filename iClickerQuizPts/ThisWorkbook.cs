﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using iClickerQuizPts.AppExceptions;


/*
 * Range names...
 * Wbk scope:
 * tblDblDippers
 * tblFirstQuizDts
 * tblQuizPts
 * ptrSemester
 * ptrCourse
 * 
 * 
 * Sheet1 scope:
 * rowSessionNmbr
 * rowCourseWk
 * rowSession
 * rowTtlPts
 */
 
 
namespace iClickerQuizPts
{
    public enum WkSession : byte
    { None = 0, First, Second, Third }

    public partial class ThisWorkbook
    {
        private bool _virginWbk = false;
        private QuizUserControl _ctrl = new QuizUserControl();
        private List<DateTime> _qDts = new List<DateTime>();
        private Excel.ListObject _tblQuizGrades = null;

        public List<DateTime> QuizDates
        {
            get
            { return _qDts; }
        }

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            this.ActionsPane.Controls.Add(_ctrl);
            this.Open += ThisWorkbook_Open;
        }

        private void ThisWorkbook_Open()
        {
            GetWbkOnOpenInfo();
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
            // Comment...
        }

        /// <summary>
        /// Method public to allow unit testing.
        /// </summary>
        public void GetWbkOnOpenInfo()
        {

        }

        public void SetListObjects()
        {
            Excel.ListObject lo;

            // First test for existence of list objects...

            // Sheet1...
            if (Globals.Sheet1.ListObjects.Count == 0)
            {
                throw new MissingNamedRangeException(
                    $"{Globals.Sheet1.Name} worksheet has no defined tables.");
            }

            // Now set listobject field, throwing an exception if we cannot find 
            // the listobject...
            lo = null;
            for(byte i = 1;i <= Globals.Sheet1.ListObjects.Count; i++ )
            {
                lo = Globals.Sheet1.ListObjects[i];
                if(lo.Name == "tblFirstQuizDts")
                {
                    _tblQuizGrades = lo;
                    break;
                }
            }
            if (_tblQuizGrades == null)
                throw new MissingNamedRangeException(
                    $"Cannot find the table \"tblFirstQuizDts\" in the {Globals.Sheet1.Name} worksheet.");
        }

        public void SetVirginWbkFlag()
        {

        }

        private void PopulateQuizDates()
        {
            DateTime dt;
            
            Excel.Range hdrs = _tblQuizGrades.HeaderRowRange;
            QuizDates.Clear();
            foreach (Excel.Range c in hdrs)
            {
                if (DateTime.TryParse(c.Value, out dt))
                    QuizDates.Add(dt);
            }
        }



        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
