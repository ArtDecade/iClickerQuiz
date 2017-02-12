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
        private DataTable _dtAllScores;
        private DataTable _dtSessions;
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
            { return _dtAllScores; }
        }

        /// <summary>
        /// Gets the <see cref="System.Data.DataTable"/> holding the session number info 
        /// from the raw iClicker data file.
        /// </summary>
        public DataTable SessionNmbrsDataTable
        {
            get
            { return _dtSessions; }
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
        public EPPlusManager(string wbkFullNm)
        {
            if (wbkFullNm.EndsWith("xlsx"))
            {
                _wbkFullNm = wbkFullNm;
                try
                {
                    ReadAppConfigDataIntoFields();
                }
                catch (InalidAppConfigItemException ex)
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
        /// Utilizes EPPlus to create two <see cref="System.Data.DataTable"/>s.  One 
        /// comprises all the data from the quiz data worksheet, the other comprises 
        /// Session number information.
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
        public virtual void CreateDataTables()
        {
            _dtAllScores = new DataTable("RawQuizData");
            _dtSessions = new DataTable("SessionHdrs");
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
                     * CREATE COLUMNS & ADD TO DATATABLES...
                     * 
                     */

                    // ALL SCORES DATA TABLE:
                    // Create a primary key column...
                    DataColumn colDataID = new DataColumn(_colNmID, typeof(int));
                    colDataID.AllowDBNull = false;
                    colDataID.AutoIncrement = true;
                    colDataID.ReadOnly = true;
                    colDataID.Unique = true;
                    _dtAllScores.Columns.Add(colDataID);

                    // Create Student ID (email) column...
                    DataColumn colStEml = new DataColumn(_colNmEmail, typeof(string));
                    colStEml.AllowDBNull = false;
                    colStEml.ReadOnly = true;
                    colStEml.Unique = true;
                    _dtAllScores.Columns.Add(colStEml);

                    // Create student Last Name column...
                    DataColumn colLn = new DataColumn(_colNmLastNm, typeof(string));
                    colLn.AllowDBNull = false;
                    _dtAllScores.Columns.Add(colLn);

                    // Create & add student First Name column...
                    DataColumn colFn = new DataColumn(_colNmFirstNm, typeof(string));
                    _dtAllScores.Columns.Add(colFn);


                    // SESSION NOS DATA TABLE:
                    // Create & add Session column...
                    DataColumn colSessNo = new DataColumn("SessionNo", typeof(string));
                    colSessNo.AllowDBNull = false;
                    colSessNo.Unique = true;
                    _dtSessions.Columns.Add(colSessNo);

                    // Create & add ComboBox column...
                    DataColumn colComboBoxEntry = new DataColumn("ComboBoxEntry", typeof(string));
                    _dtSessions.Columns.Add(colComboBoxEntry);

                    // Create & add ColHdr column...
                    DataColumn colDataHdr = new DataColumn("ColHdr", typeof(string));
                    _dtSessions.Columns.Add(colDataHdr);

                    // Create & add columns for quiz data...
                    for (int i = _firstDataCol; i <= _lastCol; i++)
                    {
                        string rawColHdr = ws.Cells[1, i].Value.ToString().Trim();
                        DataColumn col = new DataColumn(rawColHdr, typeof(byte));
                        try
                        {
                            _hdrParser.ExtractSessionDataFromColumnHeader(rawColHdr,
                                out sessionNo, out sessionDt, out maxPts);

                            // Add row to Session Nmbrs data table, trapping for dupe entries...
                            var sNos = from s in _dtSessions.AsEnumerable()
                                      where s.Field<string>("SessionNo") == sessionNo
                                      select s;
                            if (sNos.Count() == 0)
                            {
                                DataRow r = _dtSessions.NewRow();
                                r["SessionNo"] = sessionNo;
                                r["ComboBoxEntry"] = 
                                    string.Format($"Session {sessionNo} - {sessionDt}");
                                r["ColHdr"] = string.Format($"Session {sessionNo}");
                                _dtSessions.Rows.Add(r);
                            }
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
                            _dtAllScores.Columns.Add(col);
                        }
                        catch
                        {
                            InvalidQuizDataHeaderException ex = new InvalidQuizDataHeaderException();
                            ex.HeaderText = rawColHdr;
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
                        DataRow r = _dtAllScores.NewRow();
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
                        _dtAllScores.Rows.Add(r); // ...add row to dataTable
                    }
                }
            }
        }
        #endregion
    }
}

