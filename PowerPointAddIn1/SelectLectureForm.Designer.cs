namespace PowerPointAddIn1
{
    partial class SelectLectureForm
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
            this.select_lectures_combo = new System.Windows.Forms.ComboBox();
            this.select_lecture_button = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // select_lectures_combo
            // 
            this.select_lectures_combo.FormattingEnabled = true;
            this.select_lectures_combo.Location = new System.Drawing.Point(88, 32);
            this.select_lectures_combo.Name = "select_lectures_combo";
            this.select_lectures_combo.Size = new System.Drawing.Size(294, 21);
            this.select_lectures_combo.TabIndex = 0;
            // 
            // select_lecture_button
            // 
            this.select_lecture_button.Location = new System.Drawing.Point(177, 68);
            this.select_lecture_button.Name = "select_lecture_button";
            this.select_lecture_button.Size = new System.Drawing.Size(75, 23);
            this.select_lecture_button.TabIndex = 1;
            this.select_lecture_button.Text = "Select";
            this.select_lecture_button.UseVisualStyleBackColor = true;
            this.select_lecture_button.Click += new System.EventHandler(this.select_lecture_button_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(25, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Lecture:";
            // 
            // SelectLectureForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(406, 105);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.select_lecture_button);
            this.Controls.Add(this.select_lectures_combo);
            this.Name = "SelectLectureForm";
            this.Text = "Select a Lecture";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox select_lectures_combo;
        private System.Windows.Forms.Button select_lecture_button;
        private System.Windows.Forms.Label label1;
    }
}