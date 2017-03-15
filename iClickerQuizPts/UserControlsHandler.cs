using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.ComponentModel;

namespace iClickerQuizPts
{
    /// <summary>
    /// Provides a mechanism for interaction with the workbook&apos;s action panel.
    /// </summary>
    public static class UserControlsHandler
    {
        #region fields
        private static byte _crsWk;
        private static WkSession _session = WkSession.None;
        private static DataTable _dtSortedSsnsAll;
        private static EPPlusManager _eppMgr;
        private static ThisWbkDataWrapper _thisWbkWrppr = new ThisWbkDataWrapper();
        private static BindingList<Session> _blAllSessns = new BindingList<Session>();
        private static BindingList<Session> _blNewSessns = new BindingList<Session>();
        #endregion

        #region Ppts
        /// <summary>
        /// Gets the <see cref="iClickerQuizPts.WkSession"/> enumeration indicating
        /// which session/recitation within a given semester week the to-be-imported column
        /// of student quiz scores represents.
        /// </summary>
        public static WkSession WhichSession
        {
            get
            { return _session; }
        }

        /// <summary>
        /// Gets the semester week in which the to-be-imported column of student
        /// quiz scores occurred.
        /// </summary>
        public static byte CourseWeek
        {
            get
            { return _crsWk; }
        }

        /// <summary>
        /// Gets or sets the date on which the to-be-imported column of student quiz
        /// scores occurred.
        /// </summary>
        public static DateTime QuizDate { get; set; } = DateTime.Parse("1/1/2016");

        /// <summary>
        /// Gets a <see cref="System.ComponentModel.BindingList{Session}"/> of 
        /// all sessions in the raw quiz file.
        /// </summary>
        public static BindingList<Session> BindingListAllSessions
        {
            get
            { return _blAllSessns; }
        }

        /// <summary>
        /// Gets a <see cref="System.ComponentModel.BindingList{Session}"/> of 
        /// data in the raw quiz file that have not yet been imported into 
        /// iCLICKERQuizPoints worksheet.
        /// </summary>
        public static BindingList<Session> BindingListNewSessions
        {
            get
            { return _blNewSessns; }
        }
        #endregion

        #region 
        /// <summary>
        /// Sets the <see cref="iClickerQuizPts.UserControlsHandler.CourseWeek"/> property.
        /// </summary>
        /// <param name="selectedWk">The week of the semester in which the to-be-imported column of student quiz
        /// scores occurred.</param>
        public static void SetCourseWeek(string selectedWk)
        {
            _crsWk = byte.Parse(selectedWk);
        }

        /// <summary>
        /// Sets the <see cref="iClickerQuizPts.WkSession"/> property.
        /// </summary>
        /// <param name="session">Which session within a semester week represented by the 
        /// to-be-imported column of data.</param>
        public static void SetSessionEnum(string session)
        {
            switch (session)
            {
                case "First":
                    _session = WkSession.First;
                    break;
                case "Second":
                    _session = WkSession.Second;
                    break;
                case "Third":
                    _session = WkSession.Third;
                    break;
                default:
                    _session = WkSession.None;
                    break;
            }
        }

        /// <summary>
        /// Fires all other methods required to import data from an Excel file of 
        /// raw iClicker student test scores.
        /// </summary>
        public static void ImportDataMaestro()
        {
            string rawDataFileFullNm;
            bool userSelectedFile;
            userSelectedFile = PromptUserToOpenQuizDataWbk(out rawDataFileFullNm);
            if (!userSelectedFile)
                return;
            // If here user selected a file...
            _eppMgr = new EPPlusManager(rawDataFileFullNm);
            _eppMgr.CreateRawQuizDataTable();

            // Set BindingList of all Sessions...
            _blAllSessns = _eppMgr.BListSessionsAll;

            // Get BindingList of existing Sessions...
            BindingList<Session> blExisting = _thisWbkWrppr.RetrieveSessions();

            // Create BindingList of new Sessions...
            var newSessns = (from Session sAll
                             in _eppMgr.BListSessionsAll
                             orderby sAll.SessionNo
                             select sAll).Except(from Session sExisting
                                                 in blExisting
                                                 select sExisting);
            foreach (Session s in newSessns)
                _blNewSessns.Add(s);
        }

        /// <summary>
        /// Prompts user to select the Excel containing latest iClick data.
        /// </summary>
        /// <param name="dataFileFullNm">An <code>out</code> parameter to 
        /// capture the name of the selected file.</param>
        /// <returns>
        /// <see langword="true"/> if the user selected a file, otherwise
        /// <see langword="false"/>.
        /// </returns>
        /// <remarks>
        /// If the user does not select a file then the <code>dataFileFullNm</code> 
        /// out parameter will be set to <see cref="string.Empty"/>.
        /// </remarks>
        private static bool PromptUserToOpenQuizDataWbk(out string dataFileFullNm)
        {
            dataFileFullNm = string.Empty; // ...in case user cxls
            bool userSelectedWbk = new bool();

            Office.FileDialog fd = Globals.ThisWorkbook.Application.get_FileDialog(
                Office.MsoFileDialogType.msoFileDialogFilePicker);
            fd.Title = "Latest iClick Results";
            fd.AllowMultiSelect = false;
            fd.Filters.Clear();
            fd.Filters.Add("Excel Files", "*.xlsx");

            // Handle user selection...
            if (fd.Show() == -1) // ...-1 == file selected; 0 == user cxled
            {
                userSelectedWbk = true;
                dataFileFullNm = fd.SelectedItems.Item(0);
            }
            return userSelectedWbk;
        }





        






        #endregion
    }
}
