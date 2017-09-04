namespace PowerPointAddIn1
{
    partial class SelectQuestionsForm
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
            System.Windows.Forms.ListViewGroup listViewGroup1 = new System.Windows.Forms.ListViewGroup("ListViewGroup", System.Windows.Forms.HorizontalAlignment.Left);
            System.Windows.Forms.ListViewGroup listViewGroup2 = new System.Windows.Forms.ListViewGroup("ListViewGroup", System.Windows.Forms.HorizontalAlignment.Left);
            this.possibleQuestionsListView = new System.Windows.Forms.ListView();
            this.Question = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader5 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.addQuestionToSlideButton = new System.Windows.Forms.Button();
            this.labelAddQuestion = new System.Windows.Forms.Label();
            this.lectureComboAddQuestion = new System.Windows.Forms.ComboBox();
            this.chapterComboAddQuestion = new System.Windows.Forms.ComboBox();
            this.surveyComboAddQuestion = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.questionsPerSlideListView = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader6 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader7 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.removeQuestionToSlideButton = new System.Windows.Forms.Button();
            this.nextSlideButton = new System.Windows.Forms.Button();
            this.previousSlideButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // possibleQuestionsListView
            // 
            this.possibleQuestionsListView.CheckBoxes = true;
            this.possibleQuestionsListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Question,
            this.columnHeader5});
            this.possibleQuestionsListView.FullRowSelect = true;
            this.possibleQuestionsListView.GridLines = true;
            listViewGroup1.Header = "ListViewGroup";
            listViewGroup1.Name = "listViewGroup1";
            this.possibleQuestionsListView.Groups.AddRange(new System.Windows.Forms.ListViewGroup[] {
            listViewGroup1});
            this.possibleQuestionsListView.Location = new System.Drawing.Point(12, 143);
            this.possibleQuestionsListView.Name = "possibleQuestionsListView";
            this.possibleQuestionsListView.ShowGroups = false;
            this.possibleQuestionsListView.Size = new System.Drawing.Size(706, 158);
            this.possibleQuestionsListView.TabIndex = 0;
            this.possibleQuestionsListView.UseCompatibleStateImageBehavior = false;
            this.possibleQuestionsListView.View = System.Windows.Forms.View.Details;
            // 
            // Question
            // 
            this.Question.Text = "Question";
            this.Question.Width = 593;
            // 
            // columnHeader5
            // 
            this.columnHeader5.Text = "multiple choice";
            this.columnHeader5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.columnHeader5.Width = 112;
            // 
            // addQuestionToSlideButton
            // 
            this.addQuestionToSlideButton.Location = new System.Drawing.Point(643, 307);
            this.addQuestionToSlideButton.Name = "addQuestionToSlideButton";
            this.addQuestionToSlideButton.Size = new System.Drawing.Size(75, 23);
            this.addQuestionToSlideButton.TabIndex = 1;
            this.addQuestionToSlideButton.Text = "Add";
            this.addQuestionToSlideButton.UseVisualStyleBackColor = true;
            this.addQuestionToSlideButton.Click += new System.EventHandler(this.addQuestion_Click);
            // 
            // labelAddQuestion
            // 
            this.labelAddQuestion.AutoSize = true;
            this.labelAddQuestion.Location = new System.Drawing.Point(9, 9);
            this.labelAddQuestion.Name = "labelAddQuestion";
            this.labelAddQuestion.Size = new System.Drawing.Size(107, 13);
            this.labelAddQuestion.TabIndex = 2;
            this.labelAddQuestion.Text = "Add a question from: ";
            // 
            // lectureComboAddQuestion
            // 
            this.lectureComboAddQuestion.FormattingEnabled = true;
            this.lectureComboAddQuestion.Location = new System.Drawing.Point(72, 40);
            this.lectureComboAddQuestion.Name = "lectureComboAddQuestion";
            this.lectureComboAddQuestion.Size = new System.Drawing.Size(434, 21);
            this.lectureComboAddQuestion.TabIndex = 3;
            this.lectureComboAddQuestion.SelectedValueChanged += new System.EventHandler(this.comboLectures_selectionChanged);
            // 
            // chapterComboAddQuestion
            // 
            this.chapterComboAddQuestion.FormattingEnabled = true;
            this.chapterComboAddQuestion.Location = new System.Drawing.Point(72, 67);
            this.chapterComboAddQuestion.Name = "chapterComboAddQuestion";
            this.chapterComboAddQuestion.Size = new System.Drawing.Size(434, 21);
            this.chapterComboAddQuestion.TabIndex = 4;
            this.chapterComboAddQuestion.SelectedValueChanged += new System.EventHandler(this.comboChapters_selectionChanged);
            // 
            // surveyComboAddQuestion
            // 
            this.surveyComboAddQuestion.FormattingEnabled = true;
            this.surveyComboAddQuestion.Location = new System.Drawing.Point(72, 94);
            this.surveyComboAddQuestion.Name = "surveyComboAddQuestion";
            this.surveyComboAddQuestion.Size = new System.Drawing.Size(434, 21);
            this.surveyComboAddQuestion.TabIndex = 5;
            this.surveyComboAddQuestion.SelectedValueChanged += new System.EventHandler(this.comboSurveys_selectionChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 40);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Lecture:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 67);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Chapter:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 97);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(43, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Survey:";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.panel1.Location = new System.Drawing.Point(12, 348);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(706, 1);
            this.panel1.TabIndex = 9;
            // 
            // questionsPerSlideListView
            // 
            this.questionsPerSlideListView.CheckBoxes = true;
            this.questionsPerSlideListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader6,
            this.columnHeader2,
            this.columnHeader3,
            this.columnHeader7});
            this.questionsPerSlideListView.FullRowSelect = true;
            this.questionsPerSlideListView.GridLines = true;
            listViewGroup2.Header = "ListViewGroup";
            listViewGroup2.Name = "listViewGroup1";
            this.questionsPerSlideListView.Groups.AddRange(new System.Windows.Forms.ListViewGroup[] {
            listViewGroup2});
            this.questionsPerSlideListView.Location = new System.Drawing.Point(12, 379);
            this.questionsPerSlideListView.Name = "questionsPerSlideListView";
            this.questionsPerSlideListView.ShowGroups = false;
            this.questionsPerSlideListView.Size = new System.Drawing.Size(706, 180);
            this.questionsPerSlideListView.TabIndex = 10;
            this.questionsPerSlideListView.UseCompatibleStateImageBehavior = false;
            this.questionsPerSlideListView.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Question";
            this.columnHeader1.Width = 413;
            // 
            // columnHeader6
            // 
            this.columnHeader6.Text = "MC";
            this.columnHeader6.Width = 46;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Lecture";
            this.columnHeader2.Width = 100;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "Chapter";
            this.columnHeader3.Width = 93;
            // 
            // columnHeader7
            // 
            this.columnHeader7.Text = "Survey";
            this.columnHeader7.Width = 100;
            // 
            // removeQuestionToSlideButton
            // 
            this.removeQuestionToSlideButton.Location = new System.Drawing.Point(643, 565);
            this.removeQuestionToSlideButton.Name = "removeQuestionToSlideButton";
            this.removeQuestionToSlideButton.Size = new System.Drawing.Size(75, 23);
            this.removeQuestionToSlideButton.TabIndex = 11;
            this.removeQuestionToSlideButton.Text = "Remove";
            this.removeQuestionToSlideButton.UseVisualStyleBackColor = true;
            this.removeQuestionToSlideButton.Click += new System.EventHandler(this.removeQuestionsButton_Click);
            // 
            // nextSlideButton
            // 
            this.nextSlideButton.BackgroundImage = global::PowerPointAddIn1.Properties.Resources.forward;
            this.nextSlideButton.Location = new System.Drawing.Point(709, 9);
            this.nextSlideButton.Name = "nextSlideButton";
            this.nextSlideButton.Size = new System.Drawing.Size(34, 31);
            this.nextSlideButton.TabIndex = 13;
            this.nextSlideButton.Text = " ";
            this.nextSlideButton.UseVisualStyleBackColor = true;
            this.nextSlideButton.Click += new System.EventHandler(this.nextSlideButton_Click);
            // 
            // previousSlideButton
            // 
            this.previousSlideButton.BackgroundImage = global::PowerPointAddIn1.Properties.Resources.backward;
            this.previousSlideButton.Location = new System.Drawing.Point(668, 9);
            this.previousSlideButton.Name = "previousSlideButton";
            this.previousSlideButton.Size = new System.Drawing.Size(35, 31);
            this.previousSlideButton.TabIndex = 14;
            this.previousSlideButton.Text = " ";
            this.previousSlideButton.UseVisualStyleBackColor = true;
            this.previousSlideButton.Click += new System.EventHandler(this.previousSlideButton_Click);
            // 
            // SelectQuestionsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(755, 621);
            this.Controls.Add(this.previousSlideButton);
            this.Controls.Add(this.nextSlideButton);
            this.Controls.Add(this.removeQuestionToSlideButton);
            this.Controls.Add(this.questionsPerSlideListView);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.surveyComboAddQuestion);
            this.Controls.Add(this.chapterComboAddQuestion);
            this.Controls.Add(this.lectureComboAddQuestion);
            this.Controls.Add(this.labelAddQuestion);
            this.Controls.Add(this.addQuestionToSlideButton);
            this.Controls.Add(this.possibleQuestionsListView);
            this.Name = "SelectQuestionsForm";
            this.Text = "Select a question to add to slide number ";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView possibleQuestionsListView;
        private System.Windows.Forms.Button addQuestionToSlideButton;
        private System.Windows.Forms.ColumnHeader Question;
        private System.Windows.Forms.Label labelAddQuestion;
        private System.Windows.Forms.ComboBox lectureComboAddQuestion;
        private System.Windows.Forms.ComboBox chapterComboAddQuestion;
        private System.Windows.Forms.ComboBox surveyComboAddQuestion;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ListView questionsPerSlideListView;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ColumnHeader columnHeader7;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Button removeQuestionToSlideButton;
        private System.Windows.Forms.ColumnHeader columnHeader5;
        private System.Windows.Forms.Button nextSlideButton;
        private System.Windows.Forms.Button previousSlideButton;
        private System.Windows.Forms.ColumnHeader columnHeader6;
    }
}