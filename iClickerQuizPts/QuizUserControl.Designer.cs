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
            System.Windows.Forms.Label lblImportDatesCombo;
            System.Windows.Forms.Label lblCourseWkCombo;
            System.Windows.Forms.Label lblLectureSessionCombo;
            System.Windows.Forms.Label lblDate;
            System.Windows.Forms.Label lblSessions;
            this.comboCourseWeek = new System.Windows.Forms.ComboBox();
            this.comboSession = new System.Windows.Forms.ComboBox();
            this.openFileDialogQuizResults = new System.Windows.Forms.OpenFileDialog();
            this.btnOpenQuizWbk = new System.Windows.Forms.Button();
            this.comboQuizDates = new System.Windows.Forms.ComboBox();
            this.lblLatestQuizDate = new System.Windows.Forms.Label();
            this.btnImportQuizData = new System.Windows.Forms.Button();
            this.gboxDatesToShow = new System.Windows.Forms.GroupBox();
            this.radAllDates = new System.Windows.Forms.RadioButton();
            this.radNewDatesOnly = new System.Windows.Forms.RadioButton();
            this.gboxLatestQuizzes = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            lblImportDatesCombo = new System.Windows.Forms.Label();
            lblCourseWkCombo = new System.Windows.Forms.Label();
            lblLectureSessionCombo = new System.Windows.Forms.Label();
            lblDate = new System.Windows.Forms.Label();
            lblSessions = new System.Windows.Forms.Label();
            this.gboxDatesToShow.SuspendLayout();
            this.gboxLatestQuizzes.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblImportDatesCombo
            // 
            lblImportDatesCombo.AutoSize = true;
            lblImportDatesCombo.CausesValidation = false;
            lblImportDatesCombo.Enabled = false;
            lblImportDatesCombo.Location = new System.Drawing.Point(25, 214);
            lblImportDatesCombo.Name = "lblImportDatesCombo";
            lblImportDatesCombo.Size = new System.Drawing.Size(98, 13);
            lblImportDatesCombo.TabIndex = 1;
            lblImportDatesCombo.Text = "Quiz Date to Import";
            // 
            // lblCourseWkCombo
            // 
            lblCourseWkCombo.AutoSize = true;
            lblCourseWkCombo.CausesValidation = false;
            lblCourseWkCombo.Enabled = false;
            lblCourseWkCombo.Location = new System.Drawing.Point(22, 304);
            lblCourseWkCombo.Name = "lblCourseWkCombo";
            lblCourseWkCombo.Size = new System.Drawing.Size(69, 13);
            lblCourseWkCombo.TabIndex = 2;
            lblCourseWkCombo.Text = "CourseWeek";
            // 
            // lblLectureSessionCombo
            // 
            lblLectureSessionCombo.AutoSize = true;
            lblLectureSessionCombo.CausesValidation = false;
            lblLectureSessionCombo.Enabled = false;
            lblLectureSessionCombo.Location = new System.Drawing.Point(22, 391);
            lblLectureSessionCombo.Name = "lblLectureSessionCombo";
            lblLectureSessionCombo.Size = new System.Drawing.Size(83, 13);
            lblLectureSessionCombo.TabIndex = 4;
            lblLectureSessionCombo.Text = "Lecture Session";
            // 
            // lblDate
            // 
            lblDate.Enabled = false;
            lblDate.Location = new System.Drawing.Point(10, 20);
            lblDate.Name = "lblDate";
            lblDate.Size = new System.Drawing.Size(60, 13);
            lblDate.TabIndex = 12;
            lblDate.Text = "Date:";
            lblDate.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lblSessions
            // 
            lblSessions.Enabled = false;
            lblSessions.Location = new System.Drawing.Point(10, 40);
            lblSessions.Name = "lblSessions";
            lblSessions.Size = new System.Drawing.Size(60, 13);
            lblSessions.TabIndex = 13;
            lblSessions.Text = "Session(s):";
            lblSessions.TextAlign = System.Drawing.ContentAlignment.TopRight;
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
            this.comboCourseWeek.Location = new System.Drawing.Point(25, 320);
            this.comboCourseWeek.Name = "comboCourseWeek";
            this.comboCourseWeek.Size = new System.Drawing.Size(121, 21);
            this.comboCourseWeek.TabIndex = 3;
            this.comboCourseWeek.SelectedIndexChanged += new System.EventHandler(this.comboCourseWeek_SelectedIndexChanged);
            // 
            // comboSession
            // 
            this.comboSession.FormattingEnabled = true;
            this.comboSession.Items.AddRange(new object[] {
            "1st",
            "2nd",
            "3rd"});
            this.comboSession.Location = new System.Drawing.Point(22, 407);
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
            this.btnOpenQuizWbk.Location = new System.Drawing.Point(55, 95);
            this.btnOpenQuizWbk.Name = "btnOpenQuizWbk";
            this.btnOpenQuizWbk.Size = new System.Drawing.Size(139, 23);
            this.btnOpenQuizWbk.TabIndex = 6;
            this.btnOpenQuizWbk.Text = "Open Quiz File";
            this.btnOpenQuizWbk.UseVisualStyleBackColor = true;
            this.btnOpenQuizWbk.Click += new System.EventHandler(this.btnOpenQuizWbk_Click);
            // 
            // comboQuizDates
            // 
            this.comboQuizDates.FormattingEnabled = true;
            this.comboQuizDates.Location = new System.Drawing.Point(25, 230);
            this.comboQuizDates.Name = "comboQuizDates";
            this.comboQuizDates.Size = new System.Drawing.Size(200, 21);
            this.comboQuizDates.TabIndex = 7;
            this.comboQuizDates.SelectedIndexChanged += new System.EventHandler(this.comboQuizDates_SelectedIndexChanged);
            // 
            // lblLatestQuizDate
            // 
            this.lblLatestQuizDate.AutoSize = true;
            this.lblLatestQuizDate.Enabled = false;
            this.lblLatestQuizDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLatestQuizDate.Location = new System.Drawing.Point(85, 20);
            this.lblLatestQuizDate.Name = "lblLatestQuizDate";
            this.lblLatestQuizDate.Size = new System.Drawing.Size(65, 13);
            this.lblLatestQuizDate.TabIndex = 9;
            this.lblLatestQuizDate.Text = "some date";
            // 
            // btnImportQuizData
            // 
            this.btnImportQuizData.Location = new System.Drawing.Point(55, 480);
            this.btnImportQuizData.Name = "btnImportQuizData";
            this.btnImportQuizData.Size = new System.Drawing.Size(140, 23);
            this.btnImportQuizData.TabIndex = 10;
            this.btnImportQuizData.Text = "Import Quiz Data";
            this.btnImportQuizData.UseVisualStyleBackColor = true;
            this.btnImportQuizData.Click += new System.EventHandler(this.btnImportQuizData_Click);
            // 
            // gboxDatesToShow
            // 
            this.gboxDatesToShow.Controls.Add(this.radAllDates);
            this.gboxDatesToShow.Controls.Add(this.radNewDatesOnly);
            this.gboxDatesToShow.Location = new System.Drawing.Point(25, 130);
            this.gboxDatesToShow.Name = "gboxDatesToShow";
            this.gboxDatesToShow.Size = new System.Drawing.Size(200, 72);
            this.gboxDatesToShow.TabIndex = 11;
            this.gboxDatesToShow.TabStop = false;
            this.gboxDatesToShow.Text = "Dates to Show";
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
            this.radAllDates.CheckedChanged += new System.EventHandler(this.radioButton_CheckedChanged);
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
            this.radNewDatesOnly.CheckedChanged += new System.EventHandler(this.radioButton_CheckedChanged);
            // 
            // gboxLatestQuizzes
            // 
            this.gboxLatestQuizzes.BackColor = System.Drawing.SystemColors.Info;
            this.gboxLatestQuizzes.Controls.Add(this.label1);
            this.gboxLatestQuizzes.Controls.Add(lblDate);
            this.gboxLatestQuizzes.Controls.Add(lblSessions);
            this.gboxLatestQuizzes.Controls.Add(this.lblLatestQuizDate);
            this.gboxLatestQuizzes.Location = new System.Drawing.Point(25, 10);
            this.gboxLatestQuizzes.Name = "gboxLatestQuizzes";
            this.gboxLatestQuizzes.Size = new System.Drawing.Size(200, 65);
            this.gboxLatestQuizzes.TabIndex = 14;
            this.gboxLatestQuizzes.TabStop = false;
            this.gboxLatestQuizzes.Text = "Most Recent Imported Quiz(zes)";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Enabled = false;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(85, 40);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 13);
            this.label1.TabIndex = 14;
            this.label1.Text = "Numbers";
            // 
            // QuizUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.gboxLatestQuizzes);
            this.Controls.Add(this.gboxDatesToShow);
            this.Controls.Add(this.btnImportQuizData);
            this.Controls.Add(lblImportDatesCombo);
            this.Controls.Add(this.comboQuizDates);
            this.Controls.Add(this.btnOpenQuizWbk);
            this.Controls.Add(this.comboSession);
            this.Controls.Add(lblLectureSessionCombo);
            this.Controls.Add(this.comboCourseWeek);
            this.Controls.Add(lblCourseWkCombo);
            this.Location = new System.Drawing.Point(10, 0);
            this.Margin = new System.Windows.Forms.Padding(10);
            this.Name = "QuizUserControl";
            this.Padding = new System.Windows.Forms.Padding(10);
            this.Size = new System.Drawing.Size(250, 520);
            this.Load += new System.EventHandler(this.QuizUserControl_Load);
            this.gboxDatesToShow.ResumeLayout(false);
            this.gboxDatesToShow.PerformLayout();
            this.gboxLatestQuizzes.ResumeLayout(false);
            this.gboxLatestQuizzes.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ComboBox comboCourseWeek;
        private System.Windows.Forms.ComboBox comboSession;
        private System.Windows.Forms.OpenFileDialog openFileDialogQuizResults;
        private System.Windows.Forms.Button btnOpenQuizWbk;
        private System.Windows.Forms.ComboBox comboQuizDates;
        private System.Windows.Forms.Button btnImportQuizData;
        private System.Windows.Forms.GroupBox gboxDatesToShow;
        private System.Windows.Forms.RadioButton radAllDates;
        private System.Windows.Forms.RadioButton radNewDatesOnly;
        internal System.Windows.Forms.Label lblLatestQuizDate;
        private System.Windows.Forms.GroupBox gboxLatestQuizzes;
        internal System.Windows.Forms.Label label1;
    }
}
