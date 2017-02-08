using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using iClickerQuizPts.AppExceptions;

namespace iClickerQuizPts
{
    /// <summary>
    /// Provides a wrapper class for interacting with the <see cref="System.Data.DataTable"/> 
    /// of student quiz scores stored in this workbook.
    /// </summary>
    public class ThisWbkDataWrapper
    {
        // Student ID	Last Name	First Name	Semester TOTAL

        #region fields
        Excel.ListObject _loQzGrades;
        #endregion

        #region ctor
        /// <summary>
        /// Creates an instance of the <see cref="iClickerQuizPts.ThisWbkDataWrapper"/> class.
        /// </summary>
        public ThisWbkDataWrapper()
        {
            _loQzGrades = Globals.Sheet1.ListObjects["tblClkrQuizGrades"];
        }
        #endregion

        #region methods
        /// <summary>
        /// Retreives all student emails from the iCLICKERQuizPoints worksheet.
        /// </summary>
        /// <returns>All student email in the &quot;Student ID&quot; column.</returns>
        public IEnumerable<string> RetrieveStudentEmails()
        {
            Array arEmls = (Array)_loQzGrades.ListColumns["Student ID"].DataBodyRange;
            IEnumerable<string> _enumEmls = from string e in arEmls
                        orderby e
                        select e;
            return _enumEmls;
        }

        /// <summary>
        /// Retrieves the Session Numbers from the iCLICKERQuizPoints worksheet.
        /// </summary>
        /// <returns>
        /// All Session Numbers for which the worksheet has quiz scores.
        /// </returns>
        public IEnumerable<string> RetrieveSessionNumbers()
        {
            Array arColHdrs = (Array)_loQzGrades.HeaderRowRange;
            IEnumerable<string> _enumSessionNos = from string h in arColHdrs
                                                  where (h.Contains("Session"))
                                                  orderby h
                                                  select h;
            return _enumSessionNos;
        }


        #endregion



    }
}
