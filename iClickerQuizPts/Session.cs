using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Globalization;
using iClickerQuizPts.AppExceptions;

namespace iClickerQuizPts
{
    /// <summary>
    /// Represents a session during which a student takes an iClicker quiz.
    /// </summary>
    public class Session
    {
        #region fields
        private string _nmbr;
        private DateTime _date;
        private byte _maxPts;
        private string _comboBxText;
        private string _wbkColHdr;
        #endregion

        #region ppts
            #region readOnly
            /// <summary>
            /// The (unique) number of an iClicker quiz Session.
            /// </summary>
            public string SessionNo
            {
                get
                {
                    if (_nmbr.Length == 1)
                        return string.Format($"0{_nmbr}");
                    else
                        return _nmbr;
                }
            }

            /// <summary>
            /// The date of the Session.
            /// </summary>
            public DateTime QuizDate
            {
                get
                { return _date; }
            }

            /// <summary>
            /// The maximum number of points that can be earned on the iClicker 
            /// quiz given during a Session.
            /// </summary>
            public byte MaxPts
            {
                get
                { return _maxPts; }
            }

            /// <summary>
            /// Session information formatted for ComboBox display.
            /// </summary>
            /// <remarks>
            /// Property should return something like "Session 05 - 02/27/2017".
            /// </remarks>
            public string ComboBoxText
            {
                get
                {
                    string fmtdDate = _date.ToString("d", DateTimeFormatInfo.InvariantInfo);
                    return String.Format($"Session {_maxPts.ToString()} - {fmtdDate}");
                }
            }

            /// <summary>
            /// The column header to be used in the iCLICKERQuizPoints worksheet.
            /// </summary>
            public string ColHeaderText
            {
                get
                { return string.Format($"Session {_nmbr}"); }
            }
            #endregion
            #region readWrite
            /// <summary>
            /// The course week in which the Session is taught.
            /// </summary>
            public byte CourseWeek { get; set; }
            /// <summary>
            /// Which session within the course week.
            /// </summary>
            public WkSession WeeklySession { get; set; }
            #endregion
        #endregion

        #region ctors
        /// <summary>
        /// Instantiates an instance of a <see cref="iClickerQuizPts.Session"/>.
        /// </summary>
        /// <param name="rawFileHeader">
        /// The column header from a raw iClicer data file.
        /// </param>
        public Session(string rawFileHeader)
        {
            ExtractSessionDataFromColumnHeader(rawFileHeader, 
                out _nmbr, out _date, out _maxPts);
            // If necessary add a leading zero to the Session number...
            if (_nmbr.Length == 1)
                _nmbr =  string.Format($"0{_nmbr}");
        }

        /// <summary>
        /// Instantiates an instance of a <see cref="iClickerQuizPts.Session"/>.
        /// </summary>
        /// <param name="sessNo">The number of the iClicker session.</param>
        /// <param name="sessDate">The date of the session.</param>
        /// <param name="maxPts">The maximum number of points that a student 
        /// can earn from the Session&apos;s iClicker quiz.</param>
        public Session(string sessNo, DateTime sessDate, byte maxPts)
        {
            // This sessNo check SHOULD be unnecessary, but just in case...
            if (sessNo.Length == 1)
                _nmbr = string.Format($"0{sessNo}");
            else
                _nmbr = sessNo;
            _date = sessDate;
            _maxPts = maxPts;
        }
        #endregion

        #region methods
        /// <summary>
        /// Obtains the session number, quiz date, and maximum points
        /// from a raw data file data column header.
        /// </summary>
        /// <param name="hdr">A column header from a raw quiz data file.</param>
        /// <param name="sessionNo">An out parameter to capture the session number.</param>
        /// <param name="qzDate">An out parameter to capture the date of the quiz.</param>
        /// <param name="maxPts">An out parameter to capture the maximum points for the quiz.</param>
        private void ExtractSessionDataFromColumnHeader(string hdr,
            out string sessionNo, out DateTime qzDate, out byte maxPts)
        {
            try
            {
                hdr = hdr.Remove(1, "Session".Length + 1);
                hdr = hdr.Replace("Total ", string.Empty);
                // (char)91 = opening bracket (i.e., "[")...
                hdr = hdr.Replace(((char)91).ToString(), string.Empty);
                // (char)93 = closing bracket (i.e., "]")...
                hdr = hdr.Replace(((char)93).ToString(), string.Empty);
                hdr = hdr.Trim();
                // Hdr will now be something like:  "40 5/2/16 2.00"...
                int posSpace1 = hdr.IndexOf((char)34, 1); // ...(char)34 = space (i.e., " ")
                int posSpace2 = hdr.IndexOf((char)34, posSpace1 + 1);

                // Now extract our values...
                sessionNo = hdr.Substring(0, posSpace1);
                if (sessionNo.Length == 1)
                    sessionNo = string.Format($"0{sessionNo}"); // ...add leading zero, if necessary
                qzDate = DateTime.Parse( hdr.Substring(posSpace1 + 1, posSpace2 - posSpace1 - 1));
                maxPts = Byte.Parse(hdr.Substring(posSpace2 + 1));
            }
            catch (InvalidQuizDataHeaderException e)
            {
                e.HeaderText = hdr;
                throw e;
            }
        }
        #endregion

    }
}
