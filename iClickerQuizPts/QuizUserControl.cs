using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Configuration;

namespace iClickerQuizPts
{
    public partial class QuizUserControl : UserControl
    {
        private List<DateTime> allQuizDts = new List<DateTime>();
        private AppSettingsReader ar = new AppSettingsReader();

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

        private void bnOpenQuizWbk_Click(object sender, EventArgs e)
        {
            bool userSelectedWbk = ExternalWbkHandler.PromptUserToOpenQuizDataWbk();
            if(userSelectedWbk)
            {
                long ttlNmbrCols;
                string[] hdrs = ExternalWbkHandler.GetQuizFileHeaders(out ttlNmbrCols);
                allQuizDts.Clear();
                allQuizDts = ExternalWbkHandler.GetQuizDatesFromHeaders(hdrs, ttlNmbrCols);
            }

            if (userSelectedWbk && UserControlsHandler.RadBtnDates !=
                UserControlsHandler.DatesToShow.NoSelection)
                PopulateQuizDatesComboBox();
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
            
        }

        public void SetControlsEnabledProperties(bool dataFileOpened)
        {
            this.btnOpenQuizWbk.Enabled = !dataFileOpened;
            this.comboDatesToImprt.Enabled = dataFileOpened;
            this.comboCourseWeek.Enabled = dataFileOpened;
            this.gboxDatesToShow.Enabled = dataFileOpened;
            this.btnImportQuizData.Enabled = !dataFileOpened;
        }

        private void gboxDatesToShow_Enter(object sender, EventArgs e)
        {

        }

        private void radDatesToShow_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton btnDts = sender as RadioButton;
            UserControlsHandler.SetDatesToShowEnum(btnDts.Name);
            // List<DateTime> qzDts = ExternalWbkHandler.GetQuizDatesFromHeaders();
            

            // TODO:  Add handler for enabling Import Data button
        }

        private void PopulateQuizDatesComboBox()
        {
            this.comboDatesToImprt.Items.Clear();
            switch (UserControlsHandler.RadBtnDates)
            {
                case (UserControlsHandler.DatesToShow.AllDates):
                    foreach (DateTime dt in allQuizDts)
                        this.comboDatesToImprt.Items.Add(dt.ToLongDateString());
                    break;
                case (UserControlsHandler.DatesToShow.NewDatesOnly):
                    var newDts =
                        (from d in allQuizDts select d).Except(
                            from dd in Globals.ThisWorkbook.LoadedQuizDates select dd).ToList<DateTime>();

                    if (newDts.Count == 0)
                    {
                        // Notify user that all quiz dates appear to have been imported already...
                        string msg1 = (string)ar.GetValue("msgNoNewDtsToImport1", typeof(string));
                        string msg2 = (string)ar.GetValue("msgNoNewDtsToImport2", typeof(string));
                        string cptn = (string)ar.GetValue("msgNoNewDtsToImportCptn", typeof(string));
                        string msg =
                            string.Format($"{msg1} + \n\n\t + {Globals.ThisWorkbook.Name} + \n\n{msg2}");
                        MessageBox.Show(msg, cptn, MessageBoxButtons.OK);
                    }
                    else
                    {
                        foreach (DateTime ndt in newDts)
                            this.comboDatesToImprt.Items.Add(ndt.ToShortDateString());
                    }
                    break;
                default:
                    break;
            }
        }

    }
}

