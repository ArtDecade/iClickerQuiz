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
            this.lblLectureSession = new System.Windows.Forms.Label();
            this.comboSession = new System.Windows.Forms.ComboBox();
            this.openFileDialogQuizResults = new System.Windows.Forms.OpenFileDialog();
            this.btnOpenQuizWbk = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.lblLatestQuizLbl = new System.Windows.Forms.Label();
            this.lblLatestQuizDate = new System.Windows.Forms.Label();
            this.btnImportQuizData = new System.Windows.Forms.Button();
            this.gboxDatesToShow = new System.Windows.Forms.GroupBox();
            this.radNewDatesOnly = new System.Windows.Forms.RadioButton();
            this.radAllDates = new System.Windows.Forms.RadioButton();
            this.gboxDatesToShow.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblCalendar
            // 
            this.lblCalendar.AutoSize = true;
            this.lblCalendar.Enabled = false;
            this.lblCalendar.Location = new System.Drawing.Point(13, 123);
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
            // lblLectureSession
            // 
            this.lblLectureSession.AutoSize = true;
            this.lblLectureSession.Enabled = false;
            this.lblLectureSession.Location = new System.Drawing.Point(13, 352);
            this.lblLectureSession.Name = "lblLectureSession";
            this.lblLectureSession.Size = new System.Drawing.Size(83, 13);
            this.lblLectureSession.TabIndex = 4;
            this.lblLectureSession.Text = "Lecture Session";
            this.lblLectureSession.Click += new System.EventHandler(this.label1_Click);
            // 
            // comboSession
            // 
            this.comboSession.FormattingEnabled = true;
            this.comboSession.Items.AddRange(new object[] {
            "1st",
            "2nd",
            "3rd"});
            this.comboSession.Location = new System.Drawing.Point(16, 368);
            this.comboSession.Name = "comboSession";
            this.comboSession.Size = new System.Drawing.Size(121, 21);
            this.comboSession.TabIndex = 5;
            this.comboSession.SelectedIndexChanged += new System.EventHandler(this.comboSession_SelectedIndexChanged);
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
            this.btnOpenQuizWbk.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(13, 140);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 21);
            this.comboBox1.TabIndex = 7;
            // 
            // lblLatestQuizLbl
            // 
            this.lblLatestQuizLbl.Enabled = false;
            this.lblLatestQuizLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLatestQuizLbl.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.lblLatestQuizLbl.Location = new System.Drawing.Point(13, 52);
            this.lblLatestQuizLbl.Name = "lblLatestQuizLbl";
            this.lblLatestQuizLbl.Size = new System.Drawing.Size(178, 34);
            this.lblLatestQuizLbl.TabIndex = 8;
            this.lblLatestQuizLbl.Text = "Date of most recent quiz in Master file:";
            // 
            // lblLatestQuizDate
            // 
            this.lblLatestQuizDate.AutoSize = true;
            this.lblLatestQuizDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLatestQuizDate.Location = new System.Drawing.Point(13, 86);
            this.lblLatestQuizDate.Name = "lblLatestQuizDate";
            this.lblLatestQuizDate.Size = new System.Drawing.Size(41, 13);
            this.lblLatestQuizDate.TabIndex = 9;
            this.lblLatestQuizDate.Text = "label2";
            this.lblLatestQuizDate.Click += new System.EventHandler(this.lblLatestQuizDate_Click);
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
            this.gboxDatesToShow.Location = new System.Drawing.Point(13, 173);
            this.gboxDatesToShow.Name = "gboxDatesToShow";
            this.gboxDatesToShow.Size = new System.Drawing.Size(200, 70);
            this.gboxDatesToShow.TabIndex = 11;
            this.gboxDatesToShow.TabStop = false;
            this.gboxDatesToShow.Text = "Dates to Show";
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
            // 
            // QuizUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.gboxDatesToShow);
            this.Controls.Add(this.btnImportQuizData);
            this.Controls.Add(this.lblLatestQuizDate);
            this.Controls.Add(this.lblLatestQuizLbl);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.btnOpenQuizWbk);
            this.Controls.Add(this.comboSession);
            this.Controls.Add(this.lblLectureSession);
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
        private System.Windows.Forms.Label lblLectureSession;
        private System.Windows.Forms.ComboBox comboSession;
        private System.Windows.Forms.OpenFileDialog openFileDialogQuizResults;
        private System.Windows.Forms.Button btnOpenQuizWbk;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label lblLatestQuizLbl;
        private System.Windows.Forms.Label lblLatestQuizDate;
        private System.Windows.Forms.Button btnImportQuizData;
        private System.Windows.Forms.GroupBox gboxDatesToShow;
        private System.Windows.Forms.RadioButton radAllDates;
        private System.Windows.Forms.RadioButton radNewDatesOnly;
    }
}
