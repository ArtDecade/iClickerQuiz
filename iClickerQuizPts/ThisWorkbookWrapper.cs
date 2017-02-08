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
        private QuizDataLOWrapper _qdLOWrppr;
        private DblDippersLOWrapper _ddsLOWrppr;
        private NamedRangeWrapper _nrWrppr = new NamedRangeWrapper();
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
                _qdLOWrppr = new QuizDataLOWrapper(quizDataLOInfo);
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }
            try
            {
                _qdLOWrppr.SetListObjAndParentWshPpts();
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }

            // Instantiate double dippers class...
            try
            {
                _ddsLOWrppr = new DblDippersLOWrapper(dblDpprsLOInfo);
            }
            catch (ApplicationException ex)
            {
                throw ex;
            }
            try
            {
                _ddsLOWrppr.SetListObjAndParentWshPpts();
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
        public virtual void VerifyWbkScopedNames(params string[] nms)
        {
            for (int i = 0; i < nms.Length; i++)
            {
                string iClkrNm = nms[i];
                if(!_nrWrppr.WorkbookScopedRangeExists(iClkrNm))
                {
                    MissingInvalidNmdRngException ex = 
                        new MissingInvalidNmdRngException(RangeScope.Wkbk, iClkrNm);
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
        public virtual void VerifyWshScopedNames(params string[] nms)
        {
            for (int i = 0; i < nms.Length; i++)
            {
                // Since this is the only sheet holding named ranges...
                string qzDataWshNm = Globals.Sheet1.Name; 
                string iClikerNm = nms[i];
                if(!_nrWrppr.WorksheetScopedRangeExists(qzDataWshNm, iClikerNm))
                {
                    MissingInvalidNmdRngException ex =
                        new MissingInvalidNmdRngException(RangeScope.Wksheet, iClikerNm, qzDataWshNm);
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
            if (!_qdLOWrppr.ListObjectHasData && !_ddsLOWrppr.ListObjectHasData)
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
