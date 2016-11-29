namespace iClickerQuizPts
{
    partial class QuizUserControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblCalendar = new System.Windows.Forms.Label();
            this.lblCourseWk = new System.Windows.Forms.Label();
            this.comboCourseWeek = new System.Windows.Forms.ComboBox();
            this.openFileDialogQuizResults = new System.Windows.Forms.OpenFileDialog();
            this.btnOpenQuizWbk = new System.Windows.Forms.Button();
            this.comboDatesToImprt = new System.Windows.Forms.ComboBox();
            this.btnImportQuizData = new System.Windows.Forms.Button();
            this.gboxDatesToShow = new System.Windows.Forms.GroupBox();
            this.radAllDates = new System.Windows.Forms.RadioButton();
            this.radNewDatesOnly = new System.Windows.Forms.RadioButton();
            this.gboxDatesToShow.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblCalendar
            // 
            this.lblCalendar.AutoSize = true;
            this.lblCalendar.Enabled = false;
            this.lblCalendar.Location = new System.Drawing.Point(10, 162);
            this.lblCalendar.Name = "lblCalendar";
            this.lblCalendar.Size = new System.Drawing.Size(98, 13);
            this.lblCalendar.TabIndex = 1;
            this.lblCalendar.Text = "Quiz Date to Import";
            this.lblCalendar.Click += new System.EventHandler(this.lblCalendar_Click);
            // 
            // lblCourseWk
            // 
            this.lblCourseWk.AutoSize = true;
            this.lblCourseWk.Enabled = false;
            this.lblCourseWk.Location = new System.Drawing.Point(13, 276);
            this.lblCourseWk.Name = "lblCourseWk";
            this.lblCourseWk.Size = new System.Drawing.Size(69, 13);
            this.lblCourseWk.TabIndex = 2;
            this.lblCourseWk.Text = "CourseWeek";
            this.lblCourseWk.Click += new System.EventHandler(this.lblCourseWk_Click);
            // 
            // comboCourseWeek
            // 
            this.comboCourseWeek.FormattingEnabled = true;
            this.comboCourseWeek.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12"});
            this.comboCourseWeek.Location = new System.Drawing.Point(16, 292);
            this.comboCourseWeek.Name = "comboCourseWeek";
            this.comboCourseWeek.Size = new System.Drawing.Size(121, 21);
            this.comboCourseWeek.TabIndex = 3;
            this.comboCourseWeek.SelectedIndexChanged += new System.EventHandler(this.comboCourseWeek_SelectedIndexChanged);
            // 
            // openFileDialogQuizResults
            // 
            this.openFileDialogQuizResults.FileName = "openFileDialog1";
            this.openFileDialogQuizResults.Filter = "Excel Workbooks|*.xls;*.xlsx";
            this.openFileDialogQuizResults.Title = "Latest Quiz File";
            // 
            // btnOpenQuizWbk
            // 
            this.btnOpenQuizWbk.Location = new System.Drawing.Point(13, 13);
            this.btnOpenQuizWbk.Name = "btnOpenQuizWbk";
            this.btnOpenQuizWbk.Size = new System.Drawing.Size(139, 23);
            this.btnOpenQuizWbk.TabIndex = 6;
            this.btnOpenQuizWbk.Text = "Open Quiz File";
            this.btnOpenQuizWbk.UseVisualStyleBackColor = true;
            this.btnOpenQuizWbk.Click += new System.EventHandler(this.bnOpenQuizWbk_Click);
            // 
            // comboDatesToImprt
            // 
            this.comboDatesToImprt.FormattingEnabled = true;
            this.comboDatesToImprt.Location = new System.Drawing.Point(10, 179);
            this.comboDatesToImprt.Name = "comboDatesToImprt";
            this.comboDatesToImprt.Size = new System.Drawing.Size(121, 21);
            this.comboDatesToImprt.TabIndex = 7;
            // 
            // btnImportQuizData
            // 
            this.btnImportQuizData.Location = new System.Drawing.Point(16, 430);
            this.btnImportQuizData.Name = "btnImportQuizData";
            this.btnImportQuizData.Size = new System.Drawing.Size(139, 23);
            this.btnImportQuizData.TabIndex = 10;
            this.btnImportQuizData.Text = "Import Quiz Data";
            this.btnImportQuizData.UseVisualStyleBackColor = true;
            // 
            // gboxDatesToShow
            // 
            this.gboxDatesToShow.Controls.Add(this.radAllDates);
            this.gboxDatesToShow.Controls.Add(this.radNewDatesOnly);
            this.gboxDatesToShow.Location = new System.Drawing.Point(13, 89);
            this.gboxDatesToShow.Name = "gboxDatesToShow";
            this.gboxDatesToShow.Size = new System.Drawing.Size(200, 70);
            this.gboxDatesToShow.TabIndex = 11;
            this.gboxDatesToShow.TabStop = false;
            this.gboxDatesToShow.Text = "Dates to Show";
            this.gboxDatesToShow.Enter += new System.EventHandler(this.gboxDatesToShow_Enter);
            // 
            // radAllDates
            // 
            this.radAllDates.AutoSize = true;
            this.radAllDates.Location = new System.Drawing.Point(7, 44);
            this.radAllDates.Name = "radAllDates";
            this.radAllDates.Size = new System.Drawing.Size(87, 17);
            this.radAllDates.TabIndex = 1;
            this.radAllDates.TabStop = true;
            this.radAllDates.Text = "All quiz dates";
            this.radAllDates.UseVisualStyleBackColor = true;
            this.radAllDates.CheckedChanged += new System.EventHandler(this.radDatesToShow_CheckedChanged);
            // 
            // radNewDatesOnly
            // 
            this.radNewDatesOnly.AutoSize = true;
            this.radNewDatesOnly.Location = new System.Drawing.Point(7, 20);
            this.radNewDatesOnly.Name = "radNewDatesOnly";
            this.radNewDatesOnly.Size = new System.Drawing.Size(120, 17);
            this.radNewDatesOnly.TabIndex = 0;
            this.radNewDatesOnly.TabStop = true;
            this.radNewDatesOnly.Text = "New quiz dates only";
            this.radNewDatesOnly.UseVisualStyleBackColor = true;
            this.radNewDatesOnly.CheckedChanged += new System.EventHandler(this.radDatesToShow_CheckedChanged);
            // 
            // QuizUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.gboxDatesToShow);
            this.Controls.Add(this.btnImportQuizData);
            this.Controls.Add(this.comboDatesToImprt);
            this.Controls.Add(this.btnOpenQuizWbk);
            this.Controls.Add(this.comboCourseWeek);
            this.Controls.Add(this.lblCourseWk);
            this.Controls.Add(this.lblCalendar);
            this.Location = new System.Drawing.Point(10, 0);
            this.Margin = new System.Windows.Forms.Padding(10);
            this.Name = "QuizUserControl";
            this.Padding = new System.Windows.Forms.Padding(10);
            this.Size = new System.Drawing.Size(226, 480);
            this.Load += new System.EventHandler(this.QuizUserControl_Load);
            this.gboxDatesToShow.ResumeLayout(false);
            this.gboxDatesToShow.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label lblCalendar;
        private System.Windows.Forms.Label lblCourseWk;
        private System.Windows.Forms.ComboBox comboCourseWeek;
        private System.Windows.Forms.OpenFileDialog openFileDialogQuizResults;
        private System.Windows.Forms.Button btnOpenQuizWbk;
        private System.Windows.Forms.ComboBox comboDatesToImprt;
        private System.Windows.Forms.Button btnImportQuizData;
        private System.Windows.Forms.GroupBox gboxDatesToShow;
        private System.Windows.Forms.RadioButton radAllDates;
        private System.Windows.Forms.RadioButton radNewDatesOnly;
    }
}
