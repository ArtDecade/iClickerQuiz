using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using System.Configuration;

namespace iClickerQuizPts
{
    public enum WkSession : byte
    { None = 0, First, Second, Third }

    public partial class ThisWorkbook
    {
        private QuizUserControl ctrl = new QuizUserControl();
        private List<DateTime> qDts = new List<DateTime>();
        private AppSettingsReader ar = new AppSettingsReader();
        private string iClickerDataNmStub;

        public List<DateTime> LoadedQuizDates
        {
            get
            { return qDts; }
        }

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            this.ActionsPane.Controls.Add(ctrl);
            Globals.ThisWorkbook.Open += ThisWorkbook_Open;
        }

        private void ThisWorkbook_Open()
        {
            PopulateQuizDates();
            // Query app.config to get the string we search for in any open XL wbks 
            // to see if that wbk is one with iClicker data...
            iClickerDataNmStub = (string)ar.GetValue("EmbeddediClickerNm", typeof(String));

            string quizDataWbkNm;
            if(QuizWbkAlreadyOpen(out quizDataWbkNm))
            {
                // Identiy quiz data wbk...
                ExternalWbkHandler.WbkTestData =
                    this.Application.Workbooks[quizDataWbkNm];
                // Disable open external wbk button...
                ctrl.SetControlsEnabledProperties(false);
            }
            
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
            // Comment...
        }

        private void PopulateQuizDates()
        {
            DateTime dt;
            Excel.ListObject tblQuizGrades = Globals.Sheet1.ListObjects["tblClkrQuizGrades"];
            Excel.Range hdrs = tblQuizGrades.HeaderRowRange;
            LoadedQuizDates.Clear();
            foreach (Excel.Range c in hdrs)
            {
                if (DateTime.TryParse(c.Value, out dt))
                    LoadedQuizDates.Add(dt);
            }
        }

        private bool QuizWbkAlreadyOpen(out string quizWbkNm)
        {
            bool alreadyOpen = false;
            quizWbkNm = string.Empty;

            if (this.Application.Workbooks.Count > 1)
            {
                DialogResult dr;
                string cptn = (string)ar.GetValue("qstnDataWbkAlreadyOpenCptn", typeof(string));
                string msg1 = (string)ar.GetValue("qstnDataWbkAlreadyOpen1", typeof(string));
                string msg2 = (string)ar.GetValue("qstnDataWbkAlreadyOpen2", typeof(string));
                foreach (Excel.Workbook wb in Application.Workbooks)
                {
                    if(wb.Name.Contains(iClickerDataNmStub) && wb.Name != this.Name)
                    {
                        string msg = string.Format($"{msg1} \n \n\t{wb.Name} \n\n{msg2}");
                        dr = MessageBox.Show(msg,cptn,MessageBoxButtons.YesNo);
                        if (dr == DialogResult.Yes)
                        {
                            alreadyOpen = true;
                            quizWbkNm = wb.Name;
                            break;
                        }
                    }
                }
            }
            return alreadyOpen;
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
