using System;
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
using iClickerQuizPts.ListObjMgmt;


/*
 * Range names...
 * Wbk scope:
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
    /// <summary>
    /// Specifies constants defining which session/recitation within the semester week the grades are from.
    /// </summary>
    public enum WkSession : byte
    {
        /// <summary>No session has been selected yet.
        /// </summary>
        None = 0,
        /// <summary>
        /// First recitation of a given week.
        /// </summary>
        First,
        /// <summary>Second recitation of a given week.
        /// </summary>
        Second,
        /// <summary>
        /// Third recitation of a given week.
        /// </summary>
        Third
    }

    /// <summary>
    /// Provides a mechanism for pairing the name of each Excel <see cref="Excel.ListObject"/> 
    /// (i.e., Table) with its parent <see cref="Excel.Worksheet"/>. </summary>
    /// <remarks>Each of the three worksheets in this workbook contains a named 
    /// <see cref="Excel.ListObject"/>.  The <see cref="ThisWorkbook.InstantiateListObjWrapperClasses"/> method 
    /// utilizes the information stored in instances of this struct in order to verify that 
    /// the basic structure of this <see cref="Excel.Workbook"/> has not been altered.</remarks>
    public struct WshListobjPair
    {
        /// <summary>
        /// Gets the name of the Excel <see cref="Excel.ListObject"/> (i.e., Table) within 
        /// one of <c>ThisWorkbook's</c> Sheet.
        /// </summary>
        public string ListObjName { get; }
        /// <summary>
        /// Gets the name of the <c>Sheet</c> holding the identified <see cref="Excel.ListObject"/>
        /// </summary>
        public string WshNm { get; set; }
        /// <summary>
        /// Gets a value indicating whether both <see cref="WshListobjPair.ListObjName"/> and
        /// <see cref="WshListobjPair.WshNm"/> properties have been populated.
        /// </summary>
        /// <remarks>This value is set in the <see cref="WshListobjPair"/> custom constructor.  
        /// It is only set to <c>true</c> if non-empty, non-null values are provided for both 
        /// <see cref="WshListobjPair.ListObjName"/> and <see cref="WshListobjPair.WshNm"/>.
        /// <para>If the structure is instantiated via its default constructor 
        /// (which should not be used) then the value 
        /// of this property will of course remain at its default value of <c>false</c>.</para> </remarks>
        public bool PptsSet { get; }
        /// <summary>
        /// Initializes a new instance of the <see cref="WshListobjPair"/> struct.
        /// </summary>
        /// <param name="listObjNm">The name of the <see cref="Excel.ListObject"/> which the
        /// paired <see cref="Excel.Worksheet"/> contains.</param>
        /// <param name="wshNm">A worksheet within this workbook.</param>
        /// <remarks>Each worksheet in this workbook contains contains (or should contain) one
        /// and only one named <see cref="Excel.ListObject"/>.</remarks>
        public WshListobjPair(string listObjNm, string wshNm) : this()
        {
            // Set structure properties...
            ListObjName = listObjNm;
            WshNm = wshNm;
            if (!string.IsNullOrEmpty(listObjNm) && !string.IsNullOrEmpty(wshNm))
                PptsSet = true;
            else
                PptsSet = false; // ...just to be certain
        }
    }

    public partial class ThisWorkbook
    {
        #region Fields
        private bool _virginWbk = false;
        private QuizUserControl _ctrl = new QuizUserControl();
        private List<DateTime> _qDts = new List<DateTime>();
        private Excel.ListObject _tblQuizGrades = null;
        private List<WshListobjPair> _listObjsByWsh = new List<WshListobjPair>();

        private QuizDataListObjMgr _qdLOMgr;
        private DblDippersListObjMgr _ddsLOMgr;
        
        #endregion

        #region Ppts
        /// <summary>
        /// Gets a generic <c>List</c> (of type <see cref="DateTime"/>) containing the dates 
        /// of all iClicker quizzes that have been loaded into this workbook.
        /// </summary>
        public List<DateTime> QuizDates
        {
            get
            { return _qDts; }
        }

        /// <summary>
        /// Gets a <see cref="iClickerQuizPts.ListObjMgmt.ListObjectManager"/>-derived class 
        /// which handles all interaction with the <see cref="Excel.ListObject"/> containing 
        /// all iClicker quiz grades.
        /// </summary>
        public QuizDataListObjMgr QuizDataXLTblMgr
        {
            get
            { return _qdLOMgr; }
        }

        #endregion
        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            this.ActionsPane.Controls.Add(_ctrl);
            this.Open += ThisWorkbook_Open;
        }

        private void ThisWorkbook_Open()
        {
            try
            {
                InstantiateListObjWrapperClasses();
            }
            catch (InvalidWshListObjPairException ex)
            {
                MsgBoxGenerator.SetInvalidWshListObjPairMsg(ex.WshListObjPair);
                MsgBoxGenerator.ShowMsg(MessageBoxButtons.OK);
                return; // ...break out of app
            }
            catch (MissingWorksheetException ex)
            {
                MsgBoxGenerator.SetMissingWshMsg(ex.WshListObjPair);
                MsgBoxGenerator.ShowMsg(MessageBoxButtons.OK);
                return; // ...break out of app
            }
            catch (MissingListObjectException ex)
            {
                MsgBoxGenerator.SetMissingListObjMsg(ex.WshListObjPair);
                MsgBoxGenerator.ShowMsg(MessageBoxButtons.OK);
                return; // ...break out of app
            }


            try
            {
                GetWbkOnOpenInfo();
            }

            catch (MissingNamedRangeException ex)
            {
                MsgBoxGenerator.SetMissingNamedRngMsg(ex.Message);
                MsgBoxGenerator.ShowMsg(MessageBoxButtons.OK);
                return; // ...terminate program execution
            }
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
            // Comment...
        }

        /// <summary>
        /// Populates a number of properties reflecting the on-open state of this workbook.
        /// </summary>
        /// <remarks>
        /// This method calls a number of other methods, all of which populate various properties 
        /// pertaining to <list type="bullet">
        /// </list>
        /// </remarks>
        public void GetWbkOnOpenInfo()
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

        /// <summary>
        /// Instantiates instances of the <see cref="iClickerQuizPts.ListObjMgmt.ListObjectManager"/>-derived 
        /// classes that will be used in this application.
        /// </summary>
        private void InstantiateListObjWrapperClasses()
        {
            // Define the wsh-ListObj pairs...
            WshListobjPair quizDataLOInfo =
                new WshListobjPair("tblClkrQuizGrades", Globals.Sheet1.Name);
            WshListobjPair dblDpprsLOInfo =
                new WshListobjPair("tblDblDippers", Globals.Sheet2.Name);

            // Instantiate the classes...
            try
            {
                _qdLOMgr = new QuizDataListObjMgr(quizDataLOInfo);
            }
            catch(ApplicationException ex)
            {
                throw ex;
            }

            try
            {
                _ddsLOMgr = new DblDippersListObjMgr(dblDpprsLOInfo);
            }
            catch(ApplicationException ex)
            {
                throw ex;
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
