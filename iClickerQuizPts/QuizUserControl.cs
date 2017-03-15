using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace iClickerQuizPts
{
    /// <summary>
    /// Respresents the workbook&apos;s action panel.
    /// </summary>
    public partial class QuizUserControl : UserControl
    {

        /// <summary>
        /// Instantiates an instance of the workbook&apos;s action panel.
        /// </summary>
        public QuizUserControl()
        {
            InitializeComponent();
            this.Load += QuizUserControl_Load;
        }

        private void QuizUserControl_Load(object sender, EventArgs e)
        {
            if (Globals.ThisWorkbook.QuizDates.Count > 1)
                this.lblLatestQuizDate.Text =
                    Globals.ThisWorkbook.QuizDates.Max().ToShortDateString();
        }

        private void btnOpenQuizWbk_Click(object sender, EventArgs e)
        {

            // Automatically select new Sessions...
            radNewDatesOnly.Checked = true;
        }

        private void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radNewDatesOnly.Checked == true)
                this.comboSession.DataSource = UserControlsHandler.BindingListNewSessions; 
            else // ...all dates
                this.comboSession.DataSource = UserControlsHandler.BindingListAllSessions; 
        }

        private void comboQuizDates_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboCourseWeek_SelectedIndexChanged(object sender, EventArgs e)
        {
            UserControlsHandler.SetCourseWeek(comboCourseWeek.SelectedItem.ToString());
        }

        private void comboSession_SelectedIndexChanged(object sender, EventArgs e)
        {
            UserControlsHandler.SetSessionEnum(comboSession.SelectedItem.ToString());
        }

        private void btnImportQuizData_Click(object sender, EventArgs e)
        {

        }

       
    }
}
