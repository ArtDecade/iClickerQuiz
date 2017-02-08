using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Configuration;
using System.Data;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using iClickerQuizPts.AppExceptions;

namespace iClickerQuizPts
{
    /// <summary>
    /// Provides a mechanism for interacting with the iClicker-generated <see cref="Excel.Workbook"/> 
    /// containing the raw quiz data.
    /// </summary>
    public class ExternalWbkWrapper
    {
        #region fields
        private readonly byte _firstDateCol; // ...to be read from App.config file
        private Excel.Workbook _wbkTestData = null;
        #endregion

        #region Ctor
        /// <summary>
        /// Instantiates a new instance of an <see cref="iClickerQuizPts.ExternalWbkWrapper"/>.
        /// </summary>
        public ExternalWbkWrapper()
        {
            //AppSettingsReader ar = new AppSettingsReader();
            //try
            //{
            //    _firstDateCol = (Byte)ar.GetValue("FirstDataCol", typeof(Byte));
            //}
            //catch(InvalidOperationException ex)
            //{

            //}
        }
        #endregion

        #region methods
        /// <summary>
        /// Prompts user to open XL wbk with latest iClick data.
        /// </summary>
        /// <returns>
        /// Returns name of opened XL workbook (string).  
        /// If user canceled out of FileDialog returns an empty string.
        /// </returns>
        public bool PromptUserToOpenQuizDataWbk()
        {
            bool userSelectedWbk = new bool();
            string testDataWbkNm = string.Empty;
            
            Office.FileDialog fd = Globals.ThisWorkbook.Application.get_FileDialog(
                Office.MsoFileDialogType.msoFileDialogOpen);
            fd.Title = "Latest iClick Results";
            fd.AllowMultiSelect = false;
            fd.Filters.Clear();
            fd.Filters.Add("Excel Files", "*.xlsx");

            // Handle user selection...
            if (fd.Show() == -1) // ...-1 == file opened; 0 == user cxled
            {
                userSelectedWbk = true;
                fd.Execute();
                testDataWbkNm = Globals.ThisWorkbook.Application.ActiveWorkbook.Name;
                _wbkTestData = Globals.ThisWorkbook.Application.Workbooks[testDataWbkNm];
            }
            return userSelectedWbk;
        }

    

       
        #endregion
    }
}
