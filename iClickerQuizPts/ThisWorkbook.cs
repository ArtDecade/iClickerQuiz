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

using System.Configuration;
using iClickerQuizPts.AppExceptions;
using iClickerQuizPts.ListObjMgmt;
 
namespace iClickerQuizPts
{
    /// <summary>
    /// Specifies constants defining which session/recitation within the semester week the grades are from.
    /// </summary>
    public enum WkSession : byte
    {
        /// <summary>
        /// No session has been selected yet.
        /// </summary>
        None = 0,
        /// <summary>
        /// First recitation of a given week.
        /// </summary>
        First,
        /// <summary>
        /// Second recitation of a given week.
        /// </summary>
        Second,
        /// <summary>
        /// Third recitation of a given week.
        /// </summary>
        Third
    }

    /// <summary>
    /// Specifies constants indicating the scope of a named range in Excel.
    /// </summary>
    /// <remarks>
    /// This enumeration is used primarily for exception handling.
    /// </remarks>
    public enum RangeScope : byte
    {
        /// <summary>
        /// Scope not specified.
        /// </summary>
        NotSpecified = 0,
        /// <summary>
        /// A named range in Excel is scoped to the workbook.
        /// </summary>
        Wkbk = 1,
        /// <summary>
        /// A named rangle is Excel is scoped to a particular worksheet.
        /// </summary>
        Wksheet =2
    }


    /// <summary>
    /// Provides a mechanism for pairing the name of each Excel <see cref="Excel.ListObject"/> 
    /// (i.e., Table) with its parent <see cref="Excel.Worksheet"/>. </summary>
    /// <remarks>Each of the three worksheets in this workbook contains a named 
    /// <see cref="Excel.ListObject"/>.  
    /// The <see cref="ThisWbkWrapper.InstantiateListObjWrapperClasses"/> method 
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
        private string[] _wbkNmdRngs = { "ptrSemester", "ptrCourse" };
        private string[] _wshNmdRngs =
            { "rowSessionNmbr", "rowCourseWk", "rowSession", "rowTtlPts" };
        private QuizUserControl _ctrl = new QuizUserControl();
        private List<DateTime> _qDts = new List<DateTime>();
        private List<string> _sessionNos = new List<string>();
        private Excel.ListObject _tblQuizGrades = null;
        private List<WshListobjPair> _listObjsByWsh = new List<WshListobjPair>();

        private ThisWbkWrapper _twbkWrapper;

        private QuizDataLOWrapper _qdLOMgr;
        private DblDippersLOWrapper _ddsLOMgr;
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
        /// Gets a <see cref="iClickerQuizPts.ListObjMgmt.XLListObjWrapper"/>-derived class 
        /// which handles all interaction with the <see cref="Excel.ListObject"/> containing 
        /// all iClicker quiz grades.
        /// </summary>
        public QuizDataLOWrapper QuizDataXLTblMgr
        {
            get
            { return _qdLOMgr; }
        }

        #endregion

        #region wbkEventHandlers
        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            this.ActionsPane.Controls.Add(_ctrl);
            this.Open += ThisWorkbook_Open;
        }

        private void ThisWorkbook_Open()
        {
            #region VerifyStucturalIntegrityWbk
            _twbkWrapper = new ThisWbkWrapper();

            try
            {
                _twbkWrapper.InstantiateListObjWrapperClasses();
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
                _twbkWrapper.VerifyWbkScopedNames(_wbkNmdRngs);
            }
            catch (MissingInvalidNmdRngException ex)
            {
                MsgBoxGenerator.SetMissingWbkNamedRangeMsg(ex.RangeName);
                MsgBoxGenerator.ShowMsg(MessageBoxButtons.OK);
                return; // ...terminate program execution
            }

            try
            {
                string[] wshXLNms = {"rowCourseWkNmbr", "rowSessionEnum",
                    "rowSessionNmbr", "rowTtlQuizPts"};
                _twbkWrapper.VerifyWshScopedNames(wshXLNms);
            }
            catch (MissingInvalidNmdRngException ex)
            {
                MsgBoxGenerator.SetMissingInvalidWshNmdRngMsg(ex.ParentWsh, ex.RangeName);
                MsgBoxGenerator.ShowMsg(MessageBoxButtons.OK);
                return; // ...terminate program execution
            }

            try
            {
                _twbkWrapper.ReadAppConfigDataIntoFields();
            }
            catch (InalidAppConfigItemException ex)
            {
                MsgBoxGenerator.SetInvalidAppConfigKeyMsg(ex.MissingKey);
                MsgBoxGenerator.ShowMsg(MessageBoxButtons.OK);
                return; // ...terminate program execution
            }
            #endregion
            _twbkWrapper.SetVirginWbkProperty();
            if (_twbkWrapper.IsVirginWbk)
            {
                _twbkWrapper.PromptUserForCourseNameAndSemester();
                _ctrl.SetLabelForMostRecentQuizDate("No data yet.");
                _ctrl.SetLabelForMostRecentSessionNos(" -- ");
            }
            else
            // TODO:  Set latest quiz date field in control panel...
            {
               
            }
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
            // Comment...
        }
        #endregion

        

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
