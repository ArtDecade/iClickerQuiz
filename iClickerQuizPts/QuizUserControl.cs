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
    public partial class QuizUserControl : UserControl
    {
        public QuizUserControl()
        {
            InitializeComponent();
            this.Load += QuizUserControl_Load;
        }

        private void dateTimePickerQuizDate_ValueChanged(object sender, EventArgs e)
        {
            //QuizDate = dateTimePickerQuizDate.selected
        }

        private void comboCourseWeek_SelectedIndexChanged(object sender, EventArgs e)
        {
            UserControlsHandler.SetCourseWeek(comboCourseWeek.SelectedItem.ToString());
        }

        private void comboSession_SelectedIndexChanged(object sender, EventArgs e)
        {
            UserControlsHandler.SetSessionEnum(comboSession.SelectedItem.ToString());
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            UserControlsHandler.ImportDataMaestro();
        }

        #region UnusedEventHandlers
        private void lblCalendar_Click(object sender, EventArgs e)
        {

        }

        private void lblCourseWk_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        #endregion

        private void QuizUserControl_Load(object sender, EventArgs e)
        {
            if(Globals.ThisWorkbook.QuizDates.Count > 1)
            this.lblLatestQuizDate.Text =
                Globals.ThisWorkbook.QuizDates.Max().ToShortDateString();
        }

        private void lblLatestQuizDate_Click(object sender, EventArgs e)
        {

        }
    }
}
