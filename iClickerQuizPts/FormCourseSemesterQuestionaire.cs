using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace iClickerQuizPts
{
    /// <summary>
    /// Represents a form for obtaining the semester and the name of the 
    /// course from the user.
    /// </summary>
    /// <remarks>
    /// This form is only presented to users for new workbooks.  The course name
    /// and semester information are then stored in two cells in the upper
    /// left-hand corner of the QuizData worksheet.
    /// </remarks>
    public partial class FormCourseSemesterQuestionaire : Form
    {
        /// <summary>
        /// Initializes an instance of the form.
        /// </summary>
        public FormCourseSemesterQuestionaire()
        {
            InitializeComponent();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            Globals.Sheet1.Range["ptrSemester"].Value = this.txtSemester.Text;
            Globals.Sheet1.Range["ptrSemester"].Locked = true;
            Globals.Sheet1.Range["ptrCourse"].Value = this.txtSemester.Text;
            Globals.Sheet1.Range["ptrCourse"].Locked = true;

            Globals.Sheet1.Protect();
        }
    }
}
