using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        #endregion

        #region Methods
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
        /// Fires all other methods required to import data from an Excel or *.csv file of 
        /// raw iClicker student test scores.
        /// </summary>
        public static void ImportDataMaestro()
        {
            
        }
        #endregion
    }
}
