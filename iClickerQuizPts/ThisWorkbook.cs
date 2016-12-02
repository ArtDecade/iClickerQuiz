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

namespace iClickerQuizPts
{
    public enum WkSession : byte
    { None = 0, First, Second, Third }

    public partial class ThisWorkbook
    {
        private QuizUserControl ctrl = new QuizUserControl();
        private List<DateTime> qDts = new List<DateTime>();

        public List<DateTime> QuizDates
        {
            get
            { return qDts; }
        }

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            this.ActionsPane.Controls.Add(ctrl);
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
            QuizDates.Clear();
            foreach (Excel.Range c in hdrs)
            {
                if (DateTime.TryParse(c.Value, out dt))
                    QuizDates.Add(dt);
            }
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
