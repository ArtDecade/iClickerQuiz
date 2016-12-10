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
    /// <summary>
    /// Specifies constants defining which session/recitation within the semester week the grades are from.
    /// </summary>
    /// <remarks>
    /// For the purposes of this program it is assumed that the course is taught in weekly segments.  
    /// It is further assumed that within each week there are three recitation in which a student 
    /// can take a quiz.  This enumeration delineates those three recitation sessions.
    /// <para>It should be noted that a student is only supposed to take each week's iClicker quiz 
    /// once.  It has been discovered, however, that some students were attending multiple recitations 
    /// within a week and taking a week's quiz more than once.  Further, there has been no mechanism 
    /// - other than manual review of the data - for identifying students who take a week's iClicker 
    /// quiz more than once.  The entire purpose of this program is to identify and filter out those 
    /// duplicate quiz scores.</para>
    /// </remarks>
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
        Third }

    /// <summary>
    /// Provides a mechanism for pairing the name of each Excel <see cref="Excel.ListObject"/> 
    /// (i.e., Table) with its parent <see cref="Excel.Worksheet"/>. </summary>
    /// <remarks>Each of the three worksheets in this workbook contains a named 
    /// <see cref="Excel.ListObject"/>.  The <see cref="ThisWorkbook.SetListObjects"/> method 
    /// utilizes the information stored in instances of this struct in order to verify that 
    /// the basic structure of this <see cref="Excel.Workbook"/> has not been altered.</remarks>
    public struct WshListobjPairs
    {
        /// <summary>
        /// Gets or sets the name of the Excel <see cref="Excel.ListObject"/> (i.e., Table) within a 
        /// <see cref="Excel.Worksheet"/>.
        /// </summary>
        public string ListObjName { get; set; }
        /// <summary>
        /// Gets or sets the <see cref="Excel.Worksheet"/> holding the identified 
        /// <see cref="Excel.ListObject"/>
        /// </summary>
        public Excel.Worksheet Wsh { get; set; }
        /// <summary>
        /// Gets a value indicating whether both <see cref="WshListobjPairs.ListObjName"/> and
        /// <see cref="WshListobjPairs.Wsh"/> properties have been populated.
        /// </summary>
        /// <remarks>This value is set in the <see cref="WshListobjPairs"/> custom constructor.  
        /// It is only set to <c>true</c> if non-empty, non-null values are provided for both 
        /// <see cref="WshListobjPairs.ListObjName"/> and <see cref="WshListobjPairs.Wsh"/>.
        /// <para>If the structure is instantiated via its default constructor then the value 
        /// of this property will of course remain at its default value of <c>false</c>.</para> </remarks>
        public bool PptsSet { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="WshListobjPairs"/> struct.
        /// </summary>
        /// <param name="listObjNm">The name of the <see cref="Excel.ListObject"/> which the
        /// paired <see cref="Excel.Worksheet"/> contains.</param>
        /// <param name="wsh">A worksheet within this workbook.</param>
        /// <remarks>Each worksheet in this workbook contains contains (or should contain) one
        /// and only one named <see cref="Excel.ListObject"/>.</remarks>
        public WshListobjPairs(string listObjNm, Excel.Worksheet wsh) : this()
        {
            // Set structure properties...
            ListObjName = listObjNm;
            Wsh = wsh;
            if (!string.IsNullOrEmpty(listObjNm) && wsh != null)
                PptsSet = true;
            else
                PptsSet = false;
        }
    }

    public partial class ThisWorkbook
    {
        #region Fields
        private bool _virginWbk = false;
        private QuizUserControl _ctrl = new QuizUserControl();
        private List<DateTime> _qDts = new List<DateTime>();
        private Excel.ListObject _tblQuizGrades = null;
        private List<WshListobjPairs> _listObjsByWsh = new List<WshListobjPairs>();
        #endregion

        /// <summary>
        /// Gets a generic <c>List</c> (of type <see cref="DateTime"/>) containing the dates 
        /// of all iClicker quizzes that have been loaded into this workbook.
        /// </summary>
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
            try
            {
                SetListObjects();
            }
            catch(MissingNamedRangeException ex)
            {
                throw ex;
            }
        }

        public void PopulateListOfWshListObjectPairs()
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
                if(lo.Name == "tblClkrQuizGrades")
                {
                    _tblQuizGrades = lo;
                    break;
                }
            }
            if (_tblQuizGrades == null)
                throw new MissingNamedRangeException(
                    $"Cannot find the table \"tblClkrQuizGrades\" in the {Globals.Sheet1.Name} worksheet.");
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
