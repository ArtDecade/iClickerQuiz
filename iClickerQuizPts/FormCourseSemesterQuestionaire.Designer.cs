namespace iClickerQuizPts
{
    partial class FormCourseSemesterQuestionaire
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtCourseNm = new System.Windows.Forms.TextBox();
            this.lblCourse = new System.Windows.Forms.Label();
            this.lblSemester = new System.Windows.Forms.Label();
            this.txtSemester = new System.Windows.Forms.TextBox();
            this.lblInstruc = new System.Windows.Forms.Label();
            this.btnOk = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtCourseNm
            // 
            this.txtCourseNm.Location = new System.Drawing.Point(25, 88);
            this.txtCourseNm.Name = "txtCourseNm";
            this.txtCourseNm.Size = new System.Drawing.Size(225, 20);
            this.txtCourseNm.TabIndex = 0;
            // 
            // lblCourse
            // 
            this.lblCourse.AutoSize = true;
            this.lblCourse.Location = new System.Drawing.Point(25, 69);
            this.lblCourse.Name = "lblCourse";
            this.lblCourse.Size = new System.Drawing.Size(72, 13);
            this.lblCourse.TabIndex = 1;
            this.lblCourse.Text = "Course name:";
            // 
            // lblSemester
            // 
            this.lblSemester.AutoSize = true;
            this.lblSemester.Location = new System.Drawing.Point(28, 125);
            this.lblSemester.Name = "lblSemester";
            this.lblSemester.Size = new System.Drawing.Size(54, 13);
            this.lblSemester.TabIndex = 3;
            this.lblSemester.Text = "Semester:";
            // 
            // txtSemester
            // 
            this.txtSemester.Location = new System.Drawing.Point(28, 144);
            this.txtSemester.Name = "txtSemester";
            this.txtSemester.Size = new System.Drawing.Size(225, 20);
            this.txtSemester.TabIndex = 2;
            // 
            // lblInstruc
            // 
            this.lblInstruc.AutoSize = true;
            this.lblInstruc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInstruc.Location = new System.Drawing.Point(28, 23);
            this.lblInstruc.Name = "lblInstruc";
            this.lblInstruc.Size = new System.Drawing.Size(185, 16);
            this.lblInstruc.TabIndex = 4;
            this.lblInstruc.Text = "Please provide the following...";
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(100, 220);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(100, 23);
            this.btnOk.TabIndex = 5;
            this.btnOk.Text = "OK";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // FormCourseSemesterQuestionaire
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.lblInstruc);
            this.Controls.Add(this.lblSemester);
            this.Controls.Add(this.txtSemester);
            this.Controls.Add(this.lblCourse);
            this.Controls.Add(this.txtCourseNm);
            this.Name = "FormCourseSemesterQuestionaire";
            this.Text = "First-Time Information";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtCourseNm;
        private System.Windows.Forms.Label lblCourse;
        private System.Windows.Forms.Label lblSemester;
        private System.Windows.Forms.TextBox txtSemester;
        private System.Windows.Forms.Label lblInstruc;
        private System.Windows.Forms.Button btnOk;
    }
}