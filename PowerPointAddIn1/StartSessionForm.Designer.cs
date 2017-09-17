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
            this.start_session_start_button = new System.Windows.Forms.Button();
            this.start_session_lectures_combo = new System.Windows.Forms.ComboBox();
            this.start_session_error = new System.Windows.Forms.Label();
            this.numeric_seconds_spent = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.numeric_seconds_spent)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(18, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(104, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Start this session for:";
            // 
            // start_session_start_button
            // 
            this.start_session_start_button.Location = new System.Drawing.Point(150, 124);
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
            this.start_session_lectures_combo.Location = new System.Drawing.Point(128, 20);
            this.start_session_lectures_combo.Name = "start_session_lectures_combo";
            this.start_session_lectures_combo.Size = new System.Drawing.Size(233, 21);
            this.start_session_lectures_combo.TabIndex = 4;
            // 
            // start_session_error
            // 
            this.start_session_error.AutoSize = true;
            this.start_session_error.ForeColor = System.Drawing.Color.Red;
            this.start_session_error.Location = new System.Drawing.Point(185, 44);
            this.start_session_error.Name = "start_session_error";
            this.start_session_error.Size = new System.Drawing.Size(114, 13);
            this.start_session_error.TabIndex = 6;
            this.start_session_error.Text = "Please select a lecture";
            this.start_session_error.Visible = false;
            // 
            // numeric_seconds_spent
            // 
            this.numeric_seconds_spent.Location = new System.Drawing.Point(326, 81);
            this.numeric_seconds_spent.Name = "numeric_seconds_spent";
            this.numeric_seconds_spent.Size = new System.Drawing.Size(35, 20);
            this.numeric_seconds_spent.TabIndex = 7;
            this.numeric_seconds_spent.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(18, 82);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(293, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Seconds spent on slide before pushing/evluating a question:";
            // 
            // StartSessionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(377, 159);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.numeric_seconds_spent);
            this.Controls.Add(this.start_session_error);
            this.Controls.Add(this.start_session_lectures_combo);
            this.Controls.Add(this.start_session_start_button);
            this.Controls.Add(this.label1);
            this.Name = "StartSessionForm";
            this.Text = "StartSessionForm";
            this.Load += new System.EventHandler(this.StartSessionForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numeric_seconds_spent)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button start_session_start_button;
        private System.Windows.Forms.ComboBox start_session_lectures_combo;
        private System.Windows.Forms.Label start_session_error;
        private System.Windows.Forms.NumericUpDown numeric_seconds_spent;
        private System.Windows.Forms.Label label3;
    }
}