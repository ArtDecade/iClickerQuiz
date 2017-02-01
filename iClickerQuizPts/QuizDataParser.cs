using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using iClickerQuizPts.AppExceptions;

namespace iClickerQuizPts
{
    /// <summary>
    /// Provides methods for teasing out required information from
    /// a raw iClicker quiz data file.
    /// </summary>
    public class QuizDataParser
    {
        /// <summary>
        /// Obtains the session number from a quiz data column header.
        /// </summary>
        /// <param name="hdr">A column header from a raw quiz data file.</param>
        /// <param name="sessionNo">An out parameter to capture the session number.</param>
        /// <param name="qzDate">An out parameter to capture the date of the quiz.</param>
        /// <param name="maxPts">An out parameter to capture the maximum points for the quiz.</param>
        public void ExtractSessionDataFromColumnHeader(string hdr, 
            out string sessionNo, out string qzDate, out string maxPts)
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
                int space1 = hdr.IndexOf((char)34, 1); // ...(char)34 = space (i.e., " ")
                int space2 = hdr.IndexOf((char)34, space1 + 1);

                // Now extract our values...
                sessionNo = hdr.Substring(0, space1);
                qzDate = hdr.Substring(space1 + 1, space2 - space1 - 1);
                maxPts = hdr.Substring(space2 + 1);
            }
            catch (InvalidQuizDataHeaderException e)
            {
                e.HeaderText = hdr;
                throw e;
            }
        }

        /// <summary>
        /// Returns &quot;Doe&quot; given &quot;Doe, John&quot; or given
        /// simply &quot;Doe&quot;
        /// </summary>
        /// <param name="fullNm">The student&apos;s full name, in 
        /// either &quot;Last Name, First Name" or simply &quot;Last Name" format.</param>
        /// <returns>The student&apos; first name.</returns>
        public string ExtractFirstNameFromFullName(string fullNm)
        {
            string fn = string.Empty;
            int cPos;
            if (fullNm.Contains((char)44)) // ...(char)44 = comma (i.e., ",")
            {
                cPos = fullNm.IndexOf((char)44);
                fn = fullNm.Substring(cPos + 1).Trim();
            }
            return fn;
        }

        /// <summary>
        /// Returns &quot;John&quot; given &quot;Doe, John&quot; or 
        /// <see cref="string.Empty"/> given &quot;Doe&quot;
        /// simply &quot;Doe&quot;
        /// </summary>
        /// <param name="fullNm">The student&apos;s full name, in 
        /// either &quot;Last Name, First Name" or simply &quot;Last Name" format.</param>
        /// <returns>The student&apos; last name.</returns>
        public string ExtractLastNameFromFullName(string fullNm)
        {
            string ln = fullNm.Trim();
            int cPos;
            if (fullNm.Contains((char)44))
            {
                cPos = fullNm.IndexOf((char)44);
                ln = ln.Substring(0, cPos);
            }
            return ln;
        }
    }
}
