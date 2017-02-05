using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Configuration;
using iClickerQuizPts.ListObjMgmt;
using iClickerQuizPts.AppExceptions;
using Excel = Microsoft.Office.Interop.Excel;

namespace iClickerQuizPts
{
    /// <summary>
    /// Provides a mechanism for interacting with this workbook in a 
    /// unit-testable manner.
    /// </summary>
    public class ThisWorkbookWrapper
    {
        #region fields
        private bool _virginWbk;
        private byte _nmbNonScoreCols;
        private QuizDataListObjMgr _qdLOMgr;
        private DblDippersListObjMgr _ddsLOMgr;

        private NamedRangeManager _nrMgr = new NamedRangeManager();
        private string[] _wbkNmdRngs = { "ptrSemester", "ptrCourse" };
        private string[] _wshNmdRngs =
            { "rowSessionNmbr", "rowCourseWk", "rowSession", "rowTtlPts" };
        #endregion

        #region ppts
        //// QuizDataListObjMgr QuizData

        /// <summary>
        /// Gets a value indicating whether this workbook is yet populated 
        /// with any student data.
        /// </summary>
        public bool IsVirginWbk
        {
            get
            { return _virginWbk; }
        }
        #endregion

        #region methods
        /// <summary>
        /// Instantiates the (currently) 2 fields of List Object wrapper classes.
        /// </summary>
        public virtual void InstantiateListObjWrapperClasses()
        {
            // Define the wsh-ListObj pairs...
            WshListobjPair quizDataLOInfo =
                new WshListobjPair("tblClkrQuizGrades", Globals.Sheet1.Name);
            WshListobjPair dblDpprsLOInfo =
                new WshListobjPair("tblDblDippers", Globals.Sheet2.Name);

            // Instantiate quiz qata class...
            try
            {
                _qdLOMgr = new QuizDataListObjMgr(quizDataLOInfo);
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }
            try
            {
                _qdLOMgr.SetListObjAndParentWshPpts();
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }

            // Instantiate double dippers class...
            try
            {
                _ddsLOMgr = new DblDippersListObjMgr(dblDpprsLOInfo);
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }
            try
            {
                _ddsLOMgr.SetListObjAndParentWshPpts();
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Verifies that workbook-scoped named ranges exist, and that they have 
        /// valid range references.
        /// </summary>
        /// <exception cref="iClickerQuizPts.AppExceptions.MissingInvalidNmdRngException">
        /// Caught and rethrown when there are problems with the validity of a 
        /// workbook-scoped named range.</exception>
        public virtual void VerifyWbkScopedNames()
        {
            for (int i = 0; i < _wbkNmdRngs.Length; i++)
            {
                string iClkrNm = _wbkNmdRngs[i];
                try
                {
                    _nrMgr.ConfirmWorkbookScopedRangeExists(iClkrNm);
                }
                catch (MissingInvalidNmdRngException ex)
                {
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Verifies that worksheet-scoped named ranges exist, and that they have 
        /// valid range references.
        /// </summary>
        /// <exception cref="iClickerQuizPts.AppExceptions.MissingInvalidNmdRngException">
        /// Caught and rethrown when there are problems with the validity of a 
        /// worksheet-scoped named range.</exception>
        public virtual void VerifyWshScopedNames()
        {
            for (int i = 0; i < _wshNmdRngs.Length; i++)
            {
                string qzDataWshNm = Globals.Sheet1.Name; // ...since this is the only sheet holding named ranges
                string iClikerNm = _wshNmdRngs[i];
                try
                {
                    _nrMgr.ConfirmWorksheetScopedRangeExists(qzDataWshNm, iClikerNm);
                }
                catch (MissingInvalidNmdRngException ex)
                {
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Populates one or more fields with values from the <code>appSettings</code> 
        /// section of the <code>App.Config</code> file.
        /// </summary>
        /// <exception cref="iClickerQuizPts.AppExceptions.InalidAppConfigItemException">
        /// Thrown if the specified key value cannot be found in the <code>App.Config</code> file.
        /// </exception>
        public virtual void ReadAppConfigDataIntoFields()
        {
            AppSettingsReader ar = new AppSettingsReader();
            try
            {
                _nmbNonScoreCols = (byte)ar.GetValue("NmbrNonScoreCols", typeof(byte));
            }
            catch
            {
                InalidAppConfigItemException ex = new InalidAppConfigItemException();
                ex.MissingKey = "NmbrNonScoreCols";
                throw ex;
            }
        }

        /// <summary>
        /// Sets the <see cref="iClickerQuizPts.ThisWorkbookWrapper.IsVirginWbk"/> 
        /// property.
        /// </summary>
        /// <remarks>
        /// This method checks the <code>ListObjectHasData</code> property of each 
        /// <see cref="Excel.ListObject"/> in the workbook.
        /// </remarks>
        public virtual void SetVirginWbkProperty()
        {
            if (!_qdLOMgr.ListObjectHasData && !_ddsLOMgr.ListObjectHasData)
                _virginWbk = true;
        }

        /// <summary>
        /// Displays a <see cref="iClickerQuizPts.FormCourseSemesterQuestionaire"/> to 
        /// the user.
        /// </summary>
        /// <remarks>
        /// The user's input is stored in cells in the upper left-hand portion of 
        /// the <code>iCLICKERQuizPoints</code> worksheet.
        /// <para><b>NOTE:</b>&#8194;The user is only shown this form if and when 
        /// the <see cref="iClickerQuizPts.ThisWorkbookWrapper.IsVirginWbk"/> 
        /// property is <see langword="true"/>.</para>
        /// </remarks>
        public virtual void PromptUserForCourseNameAndSemester()
        {
            FormCourseSemesterQuestionaire frm = new FormCourseSemesterQuestionaire();
            frm.ShowDialog();
        }

        #endregion


    }
}
