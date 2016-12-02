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
        public WkSession WhichSession { get; set; } = WkSession.None;
        public byte CourseWeek { get; set; }
        public DateTime QuizDate { get; set; } = DateTime.Parse("1/1/2016");

        public QuizUserControl()
        {
            InitializeComponent();
        }

        
        private void dateTimePickerQuizDate_ValueChanged(object sender, EventArgs e)
        {
            QuizDate = dateTimePickerQuizDate.Value.Date;
        }

        private void comboCourseWeek_SelectedIndexChanged(object sender, EventArgs e)
        {
            CourseWeek = byte.Parse(comboCourseWeek.SelectedItem.ToString());
        }

        private void comboSession_SelectedIndexChanged(object sender, EventArgs e)
        {
            string session = comboSession.SelectedItem.ToString();
            switch(session)
            {
                case "1st":
                    WhichSession = WkSession.First;
                    break;
                case "2nd":
                    WhichSession = WkSession.Second;
                    break;
                case "3rd":
                    WhichSession = WkSession.Third;
                    break;
                default:
                    WhichSession = WkSession.None;
                    break;
            }

        }

        private void btnOK_Click(object sender, EventArgs e)
        {

        }

        private void lblCalendar_Click(object sender, EventArgs e)
        {

        }

        private void lblCourseWk_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }


    }
}
