using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ComponentModel;
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
        DataTable _dtSessNos;
        DataTable _dtEmls;
        string _mostRecentQuizDt;
        #endregion

        #region ppts
        /// <summary>
        /// Gets a <see cref="System.Data.DataTable"/> of Session 
        /// Numbers that have been loaded into this workbook.
        /// </summary>
        public DataTable SessionNmbrs
        {
            get
            { return _dtSessNos; }
        }

        /// <summary>
        /// Gets a <see cref="System.Data.DataTable"/> of email 
        /// addresses of all the students who have any quiz grade 
        /// activity at all loaded into this workbook.
        /// </summary>
        public DataTable StudentEmails
        {
            get
            { return _dtEmls; }
        }

        /// <summary>
        /// Gets the most recent quiz date within the already-imported 
        /// quiz data.
        /// </summary>
        public string MostRecentQuizDate
        {
            get
            { return _mostRecentQuizDt; }
        }

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

        /// <summary>
        /// Retreives the quiz data column headers in this workbook.
        /// </summary>
        /// <returns>All the quiz sessions already downloaded into this workbook.</returns>
        public BindingList<Session> RetrieveSessions()
        {
            int rowSessNo = Globals.Sheet1.Range["rowSessionNmbr"].Row;
            int rowDt = Globals.Sheet1.Range["xx"].Row;
            int rowMsxPts = Globals.Sheet1.Range["rowTtlQuizPts"].Row;

            BindingList<Session> bl = new BindingList<Session>();

            foreach (Excel.Range c in _loQzGrades.HeaderRowRange)
            {
                if($"{c}".Contains("Session"))
                {
                    int colNo = c.Column;
                    string sessNo = $"Globals.Sheet1.Cells[rowSessNo,colNo]";
                    Excel.Range cellDt = Globals.Sheet1.Cells[rowDt, colNo];
                    DateTime dt = DateTime.Parse(cellDt.Value);
                    Excel.Range cellPts = Globals.Sheet1.Cells[rowMsxPts, colNo];
                    Byte maxPts = byte.Parse(cellPts.Value);
                    Session s = new Session(sessNo, dt, maxPts);
                    bl.Add(s);
                }
            }
            return bl;
        }

        /// <summary>
        /// Creates and populates a <see cref="System.Data.DataTable"/> 
        /// which contains all of the Session information for quiz scores
        /// already imported into this workbook.
        /// </summary>
        public void CreateSessionNosDataTable()
        {
            int rowOffset = Globals.Sheet1.Range["rowSessionNmbr"].Row -
                _loQzGrades.HeaderRowRange.Row;

            _dtSessNos = new DataTable("ThisWbkSessionNos");
            DataColumn colSessNo = new DataColumn("SessionNo", typeof(string));
            colSessNo.AllowDBNull = false;
            colSessNo.Unique = true;
            _dtSessNos.Columns.Add(colSessNo);

            foreach(Excel.Range c in _loQzGrades.HeaderRowRange)
            {
                if(string.Format($"{c.Value}").Contains("Session"))
                {
                    DataRow r = _dtSessNos.NewRow();
                    string sNo = ((string)c.Offset[rowOffset].Value).Trim();
                    if (sNo.Length == 1)
                        sNo = "0" + sNo; // ... pad with leading zero
                    r["SessionNo"] = sNo;
                    _dtSessNos.Rows.Add(r);
                }
            }
        }

        /// <summary>
        /// Creates and populates a <see cref="System.Data.DataTable"/> 
        /// which contains all student emails already imported 
        /// into this workbook.
        /// </summary>
        public void CreateStudentEmailDataTable()
        {
            _dtEmls = new DataTable("ThisWbkEmails");
            DataColumn colEml = new DataColumn("StudentEml", typeof(string));
            colEml.AllowDBNull = false;
            colEml.Unique = true;
            _dtEmls.Columns.Add(colEml);

            Excel.Range em = 
                _loQzGrades.ListColumns["Student ID"].DataBodyRange;
            foreach(Excel.Range c in em)
            {
                DataRow r = _dtEmls.NewRow();
                r["StudentEml"] = (string)c.Value;
                _dtEmls.Rows.Add(r);
            }
        }

        /// <summary>
        /// Extracts the quiz date from a column header in the QuizPts worksheet.
        /// </summary>
        /// <param name="colHdr">A column header from the QuizPts worksheet.</param>
        /// <returns>The date of the quiz for the scores contained in the column.</returns>
        public DateTime GetDatePortionFromColHeader(string colHdr)
        {
            int posHypen = colHdr.IndexOf("-");
            return DateTime.Parse(colHdr.Substring(posHypen + 1).Trim());
        }
        #endregion
    }
}
