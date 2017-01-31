using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Configuration;
using iClickerQuizPts.AppExceptions;
using System.Data;
using System.IO;
using OfficeOpenXml;

namespace iClickerQuizPts
{
    /// <summary>
    /// Specifies constants defining the results of attempting to open an
    /// Excel file using EPPlus.
    /// </summary>
    public enum ImportResult
    {
        /// <summary>
        /// File opened successfully.
        /// </summary>
        Success = 0,
        /// <summary>
        /// The format of the data in the Excel file is incorrect.
        /// </summary>
        WrongFormat = 1,
        /// <summary>
        /// File a pre-2007 Excel file (i.e., not a *.xlsx file).
        /// </summary>
        NotExcel = 2,
        /// <summary>
        /// File cannot be imported because it is open in another application.
        /// </summary>
        StillOpen = 3
    }

    /// <summary>
    /// Provides a mechanism for utilizing EPPlus to extract data from the 
    /// Excel file containing the raw iClicker quiz points data.
    /// </summary>
    public class EPPlusManager
    {
        #region fields
        private byte _studentEmailCol;
        private byte _studentNameCol;
        private byte _firstDataCol;
        private int _lastRow;
        private int _lastCol;
        private string _wbkFullNm;
        private string _colNmID;
        private string _colNmEmail;
        private string _colNmFirstNm;
        private string _colNmLastNm;
        private DataTable _dt;
        #endregion

        #region ctor
        /// <summary>
        /// Creates an instance of the <see cref="iClickerQuizPts.EPPlusManager"/>
        /// class
        /// </summary>
        /// <param name="wbkFullNm">The full name (i.e., including path) of the
        /// Excel file containing the raw iClicker quiz points data.</param>
        /// <exception cref="iClickerQuizPts.AppExceptions.ReadingExternalWbkException">
        /// The file is either a *.csv and *.xlx files (or a different kind 
        /// of file entirely).</exception>
        public void EPPlusManger(string wbkFullNm)
        {
            if (!wbkFullNm.EndsWith("xlsx"))
            {
                ReadingExternalWbkException ex = new ReadingExternalWbkException();
                ex.ImportResult = ImportResult.NotExcel;
                throw ex;
            }
            else
                _wbkFullNm = wbkFullNm;
        }
        #endregion

        #region methods
        private void ReadAppConfigDataIntoFields()
        {
            AppSettingsReader ar = new AppSettingsReader();
            _studentEmailCol = (byte)ar.GetValue("ColNoEmailXL", typeof(byte));
            _studentNameCol = (byte)ar.GetValue("ColNoStdntNmXL", typeof(byte));
            _firstDataCol = (byte)ar.GetValue("ColNoDataBeginsXL", typeof(byte));
            _colNmID = (string)ar.GetValue("ColID", typeof(string));
            _colNmEmail = (string)ar.GetValue("ColEmail", typeof(string));
            _colNmFirstNm = (string)ar.GetValue("ColFN", typeof(string));
            _colNmLastNm = (string)ar.GetValue("ColLN", typeof(string));
        }

        /// <summary>
        /// Reads an external workbook into EPPlus.
        /// </summary>
        /// <exception cref="iClickerQuizPts.AppExceptions.ReadingExternalWbkException">
        /// There are any number of problems with the format and/or 
        /// structure of the workbook and/or the worksheet containing the quiz
        /// results data.  The exact nature of the problem is specified in
        /// the exception's message property.</exception>
        public virtual void VerifyWbkFormatIntegrityAndGetWshDimensions()
        {
            using (ExcelPackage p = new ExcelPackage())
            {
                using (FileStream stream = new FileStream(_wbkFullNm, FileMode.Open))
                {
                    // Read the workbook and it's 1st (& presumably only) worksheet...
                    p.Load(stream);
                    ExcelWorksheet ws = p.Workbook.Worksheets[1];

                    // Trap for problems with the worksheet...
                    if (ws == null)
                    {
                        ReadingExternalWbkException ex =
                            new ReadingExternalWbkException("No worksheets in the workbook.");
                        ex.ImportResult = ImportResult.WrongFormat;
                        throw ex;
                    }

                    if (!ws.Cells["A1"].Value.ToString().Trim().Equals("Student ID") ||
                        !ws.Cells["B1"].Value.ToString().Trim().Equals("Student Name") ||
                        !ws.Cells["C3"].Value.ToString().Trim().EndsWith("TOTAL"))
                    {
                        string msg = "Incorrect column headings for columns A, B, and/or C";
                        ReadingExternalWbkException ex =
                            new ReadingExternalWbkException(msg);
                        ex.ImportResult = ImportResult.WrongFormat;
                        throw ex;
                    }


                    // Find last col in wsh (header row should always have values)...
                    while (_lastCol > 1)
                    {
                        ExcelRange c = ws.Cells[1, _lastCol];
                        if (c.Value != null)
                            break;
                        else
                            _lastCol--;
                    }

                    // Trap for no data columns...
                    if (_lastCol <= _firstDataCol)
                    {
                        string msg = "There are no columns of quiz data.";
                        ReadingExternalWbkException ex = new ReadingExternalWbkException(msg);
                        ex.ImportResult = ImportResult.WrongFormat;
                        throw ex;
                    }

                    // Find last row of data in wsh.  We test the 
                    // Student ID column, which should always have an entry...
                    while (_lastRow > 1)
                    {
                        ExcelRange c = ws.Cells[_lastRow, _studentEmailCol];
                        if (c.Value != null)
                            break;
                        else
                            _lastRow--;
                    }

                    // Trap for no data rows...
                    if (_lastRow == 1)
                    {
                        string msg = "There are not data rows - there is only the header row";
                        ReadingExternalWbkException ex = new ReadingExternalWbkException(msg);
                        ex.ImportResult = ImportResult.WrongFormat;
                        throw ex;
                    }
                }
            }
        }

        /// <summary>
        /// Create and add all of the columns required to move the data from 
        /// the quiz data worksheet into a <see cref="System.Data.DataTable"/>.
        /// </summary>
        public virtual void CreateDataFreeQuizDataTable()
        {
            _dt = new DataTable("RawQuizDataTable");
            using (ExcelPackage p = new ExcelPackage())
            {
                using (FileStream stream = new FileStream(_wbkFullNm, FileMode.Open))
                {
                    // Read the workbook and it's 1st (& presumably only) worksheet...
                    p.Load(stream);
                    ExcelWorksheet ws = p.Workbook.Worksheets[1];

                    // Create a primary key column...
                    DataColumn colID = new DataColumn(_colNmID, typeof(int));
                    colID.AllowDBNull = false;
                    colID.AutoIncrement = true;
                    colID.ReadOnly = true;
                    colID.Unique = true;
                    _dt.Columns.Add(colID);

                    // Create Student ID (email) column...
                    DataColumn stID = new DataColumn(_colNmEmail, typeof(string));
                    stID.AllowDBNull = false;
                    stID.ReadOnly = true;
                    stID.Unique = true;
                    _dt.Columns.Add(stID);

                    // Create student Last Name column...
                    DataColumn ln = new DataColumn(_colNmLastNm, typeof(string));
                    ln.AllowDBNull = false;
                    _dt.Columns.Add(ln);

                    // Create & add student Firt Name column...
                    DataColumn fn = new DataColumn(_colNmFirstNm, typeof(string));
                    _dt.Columns.Add(fn);

                    // Create & add columns for quiz data...
                    for (int i = _firstDataCol; i <= _lastCol; i++)
                    {
                        string colNm = ws.Cells[1, i].Value.ToString().Trim();
                        DataColumn col = new DataColumn(colNm, typeof(byte));
                    }
                }
            }
        }

        /// <summary>
        /// Populates the <see cref="System.Data.DataTable"/> with data from
        /// the worksheet of raw quiz data.
        /// </summary>
        public virtual void PopulateTheDataTable()
        {
            using (ExcelPackage p = new ExcelPackage())
            {
                using (FileStream stream = new FileStream(_wbkFullNm, FileMode.Open))
                {
                    // Read the workbook and it's 1st (& presumably only) worksheet...
                    p.Load(stream);
                    ExcelWorksheet ws = p.Workbook.Worksheets[1];
                    string fullNm;

                    // Loop through each data row...
                    for (int rowNo = 2; rowNo <= _lastRow; rowNo++)
                    {
                        DataRow r = _dt.NewRow();
                        // Populate student name & email fields...
                        r[_colNmEmail] = ws.Cells[rowNo, _studentEmailCol].Value.ToString().Trim();
                        fullNm = ws.Cells[rowNo, _studentNameCol].Value.ToString();
                        r[_colNmLastNm] = ExtractLastNameFromFullName(fullNm);
                        r[_colNmFirstNm] = ExtractFirstNameFromFullName(fullNm);

                        // Loop through each quiz data column...
                        for (int colNo = _firstDataCol; colNo <= _lastCol; colNo++)
                        {
                            // Populate quiz data fields...
                            string colNm = ws.Cells[1, colNo].Value.ToString().Trim();
                            r[colNm] = ws.Cells[rowNo, colNo].Value;
                        }
                        _dt.Rows.Add(r); // ...add row to dataTable
                    }
                }
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
        #endregion
    }
}

