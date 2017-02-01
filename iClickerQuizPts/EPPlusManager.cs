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
        private List<string> _sessionNos = new List<string>();
        private DataTable _dt;
        private QuizDataParser _hdrParser = new QuizDataParser();
        #endregion

        #region ppts
        /// <summary>
        /// Gets the <see cref="System.Data.DataTable"/> holding the quiz scores
        /// from the raw iClicker data file.
        /// </summary>
        public DataTable RawQuizScoresDataTable
        {
            get
            { return _dt; }
        }

        /// <summary>
        /// Gets a <see cref="System.Collections.Generic.List{T}"/> of session numbers 
        /// in the raw iClicker data file.
        /// </summary>
        public List<string> SessionNos
        {
            get
            { return _sessionNos; }
        }
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
        /// <exception cref="iClickerQuizPts.AppExceptions.InalidAppConfigItemException">
        /// An entry in the <code>appSettings</code> section in the <code>App.config</code> 
        /// file could not be found.
        /// </exception>
        public void EPPlusManger(string wbkFullNm)
        {
            if (wbkFullNm.EndsWith("xlsx"))
            {
                _wbkFullNm = wbkFullNm;
                try
                {
                    ReadAppConfigDataIntoFields();
                }
                catch(InalidAppConfigItemException ex)
                {
                    throw ex;
                }
            }
            else
            {
                ReadingExternalWbkException ex = new ReadingExternalWbkException();
                ex.ImportResult = ImportResult.NotExcel;
                throw ex;
            }
        }
        #endregion

        #region methods
        private void ReadAppConfigDataIntoFields()
        {
            AppSettingsReader ar = new AppSettingsReader();
            try
            {
                _studentEmailCol = (byte)ar.GetValue("ColNoEmailXL", typeof(byte));
            }
            catch
            {
                InalidAppConfigItemException ex = new InalidAppConfigItemException();
                ex.MissingKey = "ColNoEmailXL";
                throw ex;
            }

            try
            {
                _studentNameCol = (byte)ar.GetValue("ColNoStdntNmXL", typeof(byte));
            }
            catch
            {
                InalidAppConfigItemException ex = new InalidAppConfigItemException();
                ex.MissingKey = "ColNoStdntNmXL";
                throw ex;
            }

            try
            {
                _firstDataCol = (byte)ar.GetValue("ColNoDataBeginsXL", typeof(byte));
               
            }
            catch
            {
                InalidAppConfigItemException ex = new InalidAppConfigItemException();
                ex.MissingKey = "ColNoDataBeginsXL";
                throw ex;
            }

            try
            {
                _colNmID = (string)ar.GetValue("ColID", typeof(string));
            }
            catch
            {
                InalidAppConfigItemException ex = new InalidAppConfigItemException();
                ex.MissingKey = "ColID";
                throw ex;
            }

            try
            {
                _colNmEmail = (string)ar.GetValue("ColEmail", typeof(string));
            }
            catch
            {
                InalidAppConfigItemException ex = new InalidAppConfigItemException();
                ex.MissingKey = "ColEmail";
                throw ex;
            }

            try
            {
                _colNmFirstNm = (string)ar.GetValue("ColFN", typeof(string));
            }
            catch
            {
                InalidAppConfigItemException ex = new InalidAppConfigItemException();
                ex.MissingKey = "ColFN";
                throw ex;
            }

            try
            {
                _colNmLastNm = (string)ar.GetValue("ColLN", typeof(string));
            }
            catch
            {
                InalidAppConfigItemException ex = new InalidAppConfigItemException();
                ex.MissingKey = "ColLN";
                throw ex;
            }
        }

        /// <summary>
        /// Utilizes EPPlus to move all the data from the quiz data worksheet into a 
        /// <see cref="System.Data.DataTable"/>.  
        /// <exception cref="iClickerQuizPts.AppExceptions.ReadingExternalWbkException">
        /// There are any number of problems with the format and/or 
        /// structure of the workbook and/or the worksheet containing the quiz
        /// results data.  The exact nature of the problem is specified in
        /// the exception's message property.</exception>
        /// <exception cref="iClickerQuizPts.AppExceptions.InvalidQuizDataHeaderException">
        /// A data column header in a raw iClicker data file is not in the expected 
        /// format.  As such, there are problems extracting any of:
        /// <list type="bullet">
        /// <item>session number</item>
        /// <item>quiz/session date</item>
        /// <item>maximum points for the quiz</item>
        /// </list>
        /// </exception>
        /// </summary>
        public virtual void CreateQuizScoresDataTable()
        {
            _dt = new DataTable("RawQuizDataTable");
            string sessionNo;
            string sessionDt;
            string maxPts;

            using (ExcelPackage p = new ExcelPackage())
            {
                using (FileStream stream = new FileStream(_wbkFullNm, FileMode.Open))
                {
                    // Read the workbook and it's 1st (& presumably only) worksheet...
                    p.Load(stream);
                    ExcelWorksheet ws = p.Workbook.Worksheets[1];

                    /*
                     * 
                     * TRAP FOR PROBLEMS IN THE WORKSHEET...
                     * 
                     */
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


                    /*
                     * 
                     * GATHER WORKSHEET DIMENSIONS, TRAPPING FOR MISSING DATA...
                     * 
                     */
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

                    // Find last row of data in wsh.  Student ID column should always have an entry...
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
                        string msg = "There are no rows of data.";
                        ReadingExternalWbkException ex = new ReadingExternalWbkException(msg);
                        ex.ImportResult = ImportResult.WrongFormat;
                        throw ex;
                    }


                    /*
                     * 
                     * CREATE COLUMNS & ADD TO DATATABLE...
                     * 
                     */
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
                        string colHdr = ws.Cells[1, i].Value.ToString().Trim();
                        DataColumn col = new DataColumn(colHdr, typeof(byte));
                        try
                        {
                            _hdrParser.ExtractSessionDataFromColumnHeader(colHdr,
                                out sessionNo, out sessionDt, out maxPts);

                            // Populate the List<T> off session numbers, trapping 
                            // for duplicates...
                            if(!_sessionNos.Contains(sessionNo))
                                _sessionNos.Add(sessionNo);
                            else
                            {
                                string msg = 
                                    string.Format($"Multiple instances of Session {sessionNo} are in {_wbkFullNm}.");
                                ReadingExternalWbkException ex =
                                    new ReadingExternalWbkException(msg);
                                ex.ImportResult = ImportResult.WrongFormat;
                                throw ex;
                            }

                            // Set extended properties of column, then add column...
                            col.ExtendedProperties["Session Nmbr"] = sessionNo;
                            col.ExtendedProperties["QuizDate"] = sessionDt;
                            col.ExtendedProperties["MaxQuizPts"] = maxPts;
                            col.ExtendedProperties["ComboBoxLbl"] =
                                string.Format($"Session {sessionNo} - {sessionDt}");
                            _dt.Columns.Add(col);
                        }
                        catch
                        {
                            InvalidQuizDataHeaderException ex = new InvalidQuizDataHeaderException();
                            ex.HeaderText = colHdr;
                            throw ex;
                        }
                    }


                    /*
                     * 
                     * POPULATE ROWS WITH DATA THEN ADD TO DATATABLE...
                     * 
                     */
                    string studentFullNm;

                    // Loop through each data row...
                    for (int rowNo = 2; rowNo <= _lastRow; rowNo++)
                    {
                        DataRow r = _dt.NewRow();
                        // Populate student name & email fields...
                        r[_colNmEmail] = ws.Cells[rowNo, _studentEmailCol].Value.ToString().Trim();
                        studentFullNm = ws.Cells[rowNo, _studentNameCol].Value.ToString();
                        r[_colNmLastNm] = _hdrParser.ExtractLastNameFromFullName(studentFullNm);
                        r[_colNmFirstNm] = _hdrParser.ExtractFirstNameFromFullName(studentFullNm);

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
        #endregion
    }
}

