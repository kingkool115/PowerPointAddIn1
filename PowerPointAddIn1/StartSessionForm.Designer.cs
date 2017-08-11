namespace PowerPointAddIn1
{
    partial class StartSessionForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.start_session_start_button = new System.Windows.Forms.Button();
            this.start_session_lectures_combo = new System.Windows.Forms.ComboBox();
            this.start_session_chapters_combo = new System.Windows.Forms.ComboBox();
            this.start_session_error = new System.Windows.Forms.Label();
            this.dont_record_button = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(139, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(117, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Record this session for:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(56, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Lecture:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(56, 73);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Chapter:";
            // 
            // start_session_start_button
            // 
            this.start_session_start_button.Location = new System.Drawing.Point(215, 130);
            this.start_session_start_button.Name = "start_session_start_button";
            this.start_session_start_button.Size = new System.Drawing.Size(75, 23);
            this.start_session_start_button.TabIndex = 3;
            this.start_session_start_button.Text = "Start";
            this.start_session_start_button.UseVisualStyleBackColor = true;
            this.start_session_start_button.Click += new System.EventHandler(this.start_session_start_record_button_Click);
            // 
            // start_session_lectures_combo
            // 
            this.start_session_lectures_combo.FormattingEnabled = true;
            this.start_session_lectures_combo.Location = new System.Drawing.Point(108, 38);
            this.start_session_lectures_combo.Name = "start_session_lectures_combo";
            this.start_session_lectures_combo.Size = new System.Drawing.Size(182, 21);
            this.start_session_lectures_combo.TabIndex = 4;
            this.start_session_lectures_combo.SelectionChangeCommitted += new System.EventHandler(this.lectureCombo_SelectionChanged);
            // 
            // start_session_chapters_combo
            // 
            this.start_session_chapters_combo.FormattingEnabled = true;
            this.start_session_chapters_combo.Location = new System.Drawing.Point(108, 73);
            this.start_session_chapters_combo.Name = "start_session_chapters_combo";
            this.start_session_chapters_combo.Size = new System.Drawing.Size(182, 21);
            this.start_session_chapters_combo.TabIndex = 5;
            // 
            // start_session_error
            // 
            this.start_session_error.AutoSize = true;
            this.start_session_error.ForeColor = System.Drawing.Color.Red;
            this.start_session_error.Location = new System.Drawing.Point(107, 102);
            this.start_session_error.Name = "start_session_error";
            this.start_session_error.Size = new System.Drawing.Size(183, 13);
            this.start_session_error.TabIndex = 6;
            this.start_session_error.Text = "Please select a lecture and a chapter";
            this.start_session_error.Visible = false;
            // 
            // dont_record_button
            // 
            this.dont_record_button.Location = new System.Drawing.Point(110, 130);
            this.dont_record_button.Name = "dont_record_button";
            this.dont_record_button.Size = new System.Drawing.Size(75, 23);
            this.dont_record_button.TabIndex = 7;
            this.dont_record_button.Text = "Don\'t record";
            this.dont_record_button.UseVisualStyleBackColor = true;
            this.dont_record_button.Click += new System.EventHandler(this.dontRecordButton_Click);
            // 
            // StartSessionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(377, 165);
            this.Controls.Add(this.dont_record_button);
            this.Controls.Add(this.start_session_error);
            this.Controls.Add(this.start_session_chapters_combo);
            this.Controls.Add(this.start_session_lectures_combo);
            this.Controls.Add(this.start_session_start_button);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "StartSessionForm";
            this.Text = "StartSessionForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button start_session_start_button;
        private System.Windows.Forms.ComboBox start_session_lectures_combo;
        private System.Windows.Forms.ComboBox start_session_chapters_combo;
        private System.Windows.Forms.Label start_session_error;
        private System.Windows.Forms.Button dont_record_button;
    }
}