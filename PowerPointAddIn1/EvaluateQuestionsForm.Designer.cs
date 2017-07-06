namespace PowerPointAddIn1
{
    partial class EvaluateQuestionsForm
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
            this.notEvaluatedQuestionsListView = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader11 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader5 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.label1 = new System.Windows.Forms.Label();
            this.evaluateQuestionsListView = new System.Windows.Forms.ListView();
            this.columnHeader6 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader12 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader7 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader8 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader9 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader10 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.evaluateQuestionsButton = new System.Windows.Forms.Button();
            this.removeQuestionEvaluationButton = new System.Windows.Forms.Button();
            this.previousSlideButton = new System.Windows.Forms.Button();
            this.nextSlideButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // notEvaluatedQuestionsListView
            // 
            this.notEvaluatedQuestionsListView.CheckBoxes = true;
            this.notEvaluatedQuestionsListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader11,
            this.columnHeader2,
            this.columnHeader3,
            this.columnHeader4,
            this.columnHeader5});
            this.notEvaluatedQuestionsListView.FullRowSelect = true;
            this.notEvaluatedQuestionsListView.GridLines = true;
            this.notEvaluatedQuestionsListView.Location = new System.Drawing.Point(12, 60);
            this.notEvaluatedQuestionsListView.Name = "notEvaluatedQuestionsListView";
            this.notEvaluatedQuestionsListView.Size = new System.Drawing.Size(534, 149);
            this.notEvaluatedQuestionsListView.TabIndex = 0;
            this.notEvaluatedQuestionsListView.UseCompatibleStateImageBehavior = false;
            this.notEvaluatedQuestionsListView.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Question";
            this.columnHeader1.Width = 200;
            // 
            // columnHeader11
            // 
            this.columnHeader11.Text = "pushed at slide";
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "MC";
            this.columnHeader2.Width = 30;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "Lecture";
            this.columnHeader3.Width = 100;
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "Chapter";
            this.columnHeader4.Width = 100;
            // 
            // columnHeader5
            // 
            this.columnHeader5.Text = "Survey";
            this.columnHeader5.Width = 100;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 42);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(140, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Open questions to evaluate:";
            // 
            // evaluateQuestionsListView
            // 
            this.evaluateQuestionsListView.CheckBoxes = true;
            this.evaluateQuestionsListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader6,
            this.columnHeader12,
            this.columnHeader7,
            this.columnHeader8,
            this.columnHeader9,
            this.columnHeader10});
            this.evaluateQuestionsListView.FullRowSelect = true;
            this.evaluateQuestionsListView.GridLines = true;
            this.evaluateQuestionsListView.Location = new System.Drawing.Point(12, 271);
            this.evaluateQuestionsListView.Name = "evaluateQuestionsListView";
            this.evaluateQuestionsListView.Size = new System.Drawing.Size(534, 149);
            this.evaluateQuestionsListView.TabIndex = 2;
            this.evaluateQuestionsListView.UseCompatibleStateImageBehavior = false;
            this.evaluateQuestionsListView.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader6
            // 
            this.columnHeader6.Text = "Question";
            this.columnHeader6.Width = 200;
            // 
            // columnHeader12
            // 
            this.columnHeader12.Text = "pushed at slide";
            // 
            // columnHeader7
            // 
            this.columnHeader7.Text = "MC";
            this.columnHeader7.Width = 30;
            // 
            // columnHeader8
            // 
            this.columnHeader8.Text = "Lecture";
            this.columnHeader8.Width = 100;
            // 
            // columnHeader9
            // 
            this.columnHeader9.Text = "Chapter";
            this.columnHeader9.Width = 100;
            // 
            // columnHeader10
            // 
            this.columnHeader10.Text = "Survey";
            this.columnHeader10.Width = 100;
            // 
            // evaluateQuestionsButton
            // 
            this.evaluateQuestionsButton.Location = new System.Drawing.Point(419, 215);
            this.evaluateQuestionsButton.Name = "evaluateQuestionsButton";
            this.evaluateQuestionsButton.Size = new System.Drawing.Size(127, 23);
            this.evaluateQuestionsButton.TabIndex = 3;
            this.evaluateQuestionsButton.Text = "Evaluate on this slide";
            this.evaluateQuestionsButton.UseVisualStyleBackColor = true;
            this.evaluateQuestionsButton.Click += new System.EventHandler(this.evaluateQuestionsButton_Click);
            // 
            // removeQuestionEvaluationButton
            // 
            this.removeQuestionEvaluationButton.Location = new System.Drawing.Point(370, 426);
            this.removeQuestionEvaluationButton.Name = "removeQuestionEvaluationButton";
            this.removeQuestionEvaluationButton.Size = new System.Drawing.Size(176, 23);
            this.removeQuestionEvaluationButton.TabIndex = 4;
            this.removeQuestionEvaluationButton.Text = "Remove Evaluation on this slide";
            this.removeQuestionEvaluationButton.UseVisualStyleBackColor = true;
            this.removeQuestionEvaluationButton.Click += new System.EventHandler(this.removeQuestionEvaluationButton_Click);
            // 
            // previousSlideButton
            // 
            this.previousSlideButton.BackgroundImage = global::PowerPointAddIn1.Properties.Resources.backward;
            this.previousSlideButton.Location = new System.Drawing.Point(470, 12);
            this.previousSlideButton.Name = "previousSlideButton";
            this.previousSlideButton.Size = new System.Drawing.Size(35, 31);
            this.previousSlideButton.TabIndex = 5;
            this.previousSlideButton.Text = "  ";
            this.previousSlideButton.UseVisualStyleBackColor = true;
            this.previousSlideButton.Click += new System.EventHandler(this.previousSlideButton_Click);
            // 
            // nextSlideButton
            // 
            this.nextSlideButton.BackgroundImage = global::PowerPointAddIn1.Properties.Resources.forward;
            this.nextSlideButton.Location = new System.Drawing.Point(511, 12);
            this.nextSlideButton.Name = "nextSlideButton";
            this.nextSlideButton.Size = new System.Drawing.Size(35, 31);
            this.nextSlideButton.TabIndex = 6;
            this.nextSlideButton.Text = " ";
            this.nextSlideButton.UseVisualStyleBackColor = true;
            this.nextSlideButton.Click += new System.EventHandler(this.nextSlideButton_Click);
            // 
            // EvaluateQuestionsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(558, 491);
            this.Controls.Add(this.nextSlideButton);
            this.Controls.Add(this.previousSlideButton);
            this.Controls.Add(this.removeQuestionEvaluationButton);
            this.Controls.Add(this.evaluateQuestionsButton);
            this.Controls.Add(this.evaluateQuestionsListView);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.notEvaluatedQuestionsListView);
            this.Name = "EvaluateQuestionsForm";
            this.Text = "Select a question which should be evaluated on this slide.";
            this.Load += new System.EventHandler(this.SelectAnswersForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView notEvaluatedQuestionsListView;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.ColumnHeader columnHeader5;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListView evaluateQuestionsListView;
        private System.Windows.Forms.ColumnHeader columnHeader6;
        private System.Windows.Forms.ColumnHeader columnHeader7;
        private System.Windows.Forms.ColumnHeader columnHeader8;
        private System.Windows.Forms.ColumnHeader columnHeader9;
        private System.Windows.Forms.ColumnHeader columnHeader10;
        private System.Windows.Forms.Button evaluateQuestionsButton;
        private System.Windows.Forms.Button removeQuestionEvaluationButton;
        private System.Windows.Forms.Button previousSlideButton;
        private System.Windows.Forms.Button nextSlideButton;
        private System.Windows.Forms.ColumnHeader columnHeader11;
        private System.Windows.Forms.ColumnHeader columnHeader12;
    }
}