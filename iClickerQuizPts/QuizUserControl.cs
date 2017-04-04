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

        /// <summary>
        /// Updates the label at the top of the 
        /// <see cref="iClickerQuizPts.QuizUserControl"/> to display the date of 
        /// the most recent quizz(es) in this workbook.
        /// </summary>
        /// <param name="quizDate">The date of the most recent quizz(es) that 
        /// have been loaded into this workbook.</param>
        public void SetLabelForMostRecentQuizDate(string quizDate)
        {
            this.lblLatestQuizDate.Text = quizDate;
        }

        /// <summary>
        /// Updates the label at the top of the 
        /// <see cref="iClickerQuizPts.QuizUserControl"/> to display the session 
        /// number(s) for the most recent quizz(es) in this workbook.
        /// </summary>
        /// <param name="sessNos">The Session number for the most recent quiz 
        /// that has been imported into this workbook.  If more than one quiz 
        /// was administered on that data a comma-delimited list of those 
        /// Session numbers.</param>
        public void SetLabelForMostRecentSessionNos(string sessNos)
        {
            this.lblMostRecentSessNos.Text = sessNos;
        }

       
    }
}
