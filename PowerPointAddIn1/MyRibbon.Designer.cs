namespace PowerPointAddIn1
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MyRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MyRibbon));
            this.LARS = this.Factory.CreateRibbonTab();
            this.groupCreateNewSurvey = this.Factory.CreateRibbonGroup();
            this.groupConnect = this.Factory.CreateRibbonGroup();
            this.groupSelectSurvey = this.Factory.CreateRibbonGroup();
            this.lectureDropDown = this.Factory.CreateRibbonDropDown();
            this.chapterDropDown = this.Factory.CreateRibbonDropDown();
            this.surveyDropDown = this.Factory.CreateRibbonDropDown();
            this.startSurveyGroup = this.Factory.CreateRibbonGroup();
            this.addQuestionGroup = this.Factory.CreateRibbonGroup();
            this.btnCreateNewSurvey = this.Factory.CreateRibbonButton();
            this.connectBtn = this.Factory.CreateRibbonButton();
            this.startSurveyButton = this.Factory.CreateRibbonButton();
            this.buttonAddQuestion = this.Factory.CreateRibbonButton();
            this.buttonRemoveQuestion = this.Factory.CreateRibbonButton();
            this.answerGroup = this.Factory.CreateRibbonGroup();
            this.buttonAddAnswer = this.Factory.CreateRibbonButton();
            this.buttonRemoveAnswer = this.Factory.CreateRibbonButton();
            this.LARS.SuspendLayout();
            this.groupCreateNewSurvey.SuspendLayout();
            this.groupConnect.SuspendLayout();
            this.groupSelectSurvey.SuspendLayout();
            this.startSurveyGroup.SuspendLayout();
            this.addQuestionGroup.SuspendLayout();
            this.answerGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // LARS
            // 
            this.LARS.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.LARS.Groups.Add(this.groupCreateNewSurvey);
            this.LARS.Groups.Add(this.groupConnect);
            this.LARS.Groups.Add(this.groupSelectSurvey);
            this.LARS.Groups.Add(this.startSurveyGroup);
            this.LARS.Groups.Add(this.addQuestionGroup);
            this.LARS.Groups.Add(this.answerGroup);
            this.LARS.Label = "LARS";
            this.LARS.Name = "LARS";
            // 
            // groupCreateNewSurvey
            // 
            this.groupCreateNewSurvey.Items.Add(this.btnCreateNewSurvey);
            this.groupCreateNewSurvey.Label = "  New Survey  ";
            this.groupCreateNewSurvey.Name = "groupCreateNewSurvey";
            // 
            // groupConnect
            // 
            this.groupConnect.Items.Add(this.connectBtn);
            this.groupConnect.Label = "  Not Connected  ";
            this.groupConnect.Name = "groupConnect";
            // 
            // groupSelectSurvey
            // 
            this.groupSelectSurvey.Items.Add(this.lectureDropDown);
            this.groupSelectSurvey.Items.Add(this.chapterDropDown);
            this.groupSelectSurvey.Items.Add(this.surveyDropDown);
            this.groupSelectSurvey.Label = "Select survey";
            this.groupSelectSurvey.Name = "groupSelectSurvey";
            // 
            // lectureDropDown
            // 
            this.lectureDropDown.Enabled = false;
            this.lectureDropDown.Label = "Lecture: ";
            this.lectureDropDown.Name = "lectureDropDown";
            this.lectureDropDown.SizeString = "XXXXXXXXXXXXXXXXXXXXXXXXXX";
            this.lectureDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.lectureDropDown_SelectionChanged);
            // 
            // chapterDropDown
            // 
            this.chapterDropDown.Enabled = false;
            this.chapterDropDown.Label = "Chapter: ";
            this.chapterDropDown.Name = "chapterDropDown";
            this.chapterDropDown.SizeString = "XXXXXXXXXXXXXXXXXXXXXXXXXX";
            this.chapterDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chapterDropDown_SelectionChanged);
            // 
            // surveyDropDown
            // 
            this.surveyDropDown.Enabled = false;
            this.surveyDropDown.Label = "Survey:  ";
            this.surveyDropDown.Name = "surveyDropDown";
            this.surveyDropDown.SizeString = "XXXXXXXXXXXXXXXXXXXXXXXXXX";
            // 
            // startSurveyGroup
            // 
            this.startSurveyGroup.Items.Add(this.startSurveyButton);
            this.startSurveyGroup.Label = "  Start Survey  ";
            this.startSurveyGroup.Name = "startSurveyGroup";
            // 
            // addQuestionGroup
            // 
            this.addQuestionGroup.Items.Add(this.buttonAddQuestion);
            this.addQuestionGroup.Items.Add(this.buttonRemoveQuestion);
            this.addQuestionGroup.Label = "  Question  ";
            this.addQuestionGroup.Name = "addQuestionGroup";
            // 
            // btnCreateNewSurvey
            // 
            this.btnCreateNewSurvey.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCreateNewSurvey.Image = ((System.Drawing.Image)(resources.GetObject("btnCreateNewSurvey.Image")));
            this.btnCreateNewSurvey.Label = " Create";
            this.btnCreateNewSurvey.Name = "btnCreateNewSurvey";
            this.btnCreateNewSurvey.ShowImage = true;
            this.btnCreateNewSurvey.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateNewSurvey_Click);
            // 
            // connectBtn
            // 
            this.connectBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.connectBtn.Image = ((System.Drawing.Image)(resources.GetObject("connectBtn.Image")));
            this.connectBtn.Label = " ";
            this.connectBtn.Name = "connectBtn";
            this.connectBtn.ShowImage = true;
            this.connectBtn.Tag = "connect";
            this.connectBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.connectBtn_Click);
            // 
            // startSurveyButton
            // 
            this.startSurveyButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.startSurveyButton.Image = ((System.Drawing.Image)(resources.GetObject("startSurveyButton.Image")));
            this.startSurveyButton.Label = " ";
            this.startSurveyButton.Name = "startSurveyButton";
            this.startSurveyButton.ShowImage = true;
            // 
            // buttonAddQuestion
            // 
            this.buttonAddQuestion.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAddQuestion.Enabled = false;
            this.buttonAddQuestion.Image = ((System.Drawing.Image)(resources.GetObject("buttonAddQuestion.Image")));
            this.buttonAddQuestion.Label = "  Add  ";
            this.buttonAddQuestion.Name = "buttonAddQuestion";
            this.buttonAddQuestion.ShowImage = true;
            this.buttonAddQuestion.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAddQuestion_Click);
            // 
            // buttonRemoveQuestion
            // 
            this.buttonRemoveQuestion.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonRemoveQuestion.Enabled = false;
            this.buttonRemoveQuestion.Image = global::PowerPointAddIn1.Properties.Resources.delete;
            this.buttonRemoveQuestion.Label = "  Remove  ";
            this.buttonRemoveQuestion.Name = "buttonRemoveQuestion";
            this.buttonRemoveQuestion.ShowImage = true;
            // 
            // answerGroup
            // 
            this.answerGroup.Items.Add(this.buttonAddAnswer);
            this.answerGroup.Items.Add(this.buttonRemoveAnswer);
            this.answerGroup.Label = "  Evaluation  ";
            this.answerGroup.Name = "answerGroup";
            // 
            // buttonAddAnswer
            // 
            this.buttonAddAnswer.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAddAnswer.Enabled = false;
            this.buttonAddAnswer.Image = global::PowerPointAddIn1.Properties.Resources.add_answer;
            this.buttonAddAnswer.Label = "  Show  ";
            this.buttonAddAnswer.Name = "buttonAddAnswer";
            this.buttonAddAnswer.ShowImage = true;
            // 
            // buttonRemoveAnswer
            // 
            this.buttonRemoveAnswer.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonRemoveAnswer.Enabled = false;
            this.buttonRemoveAnswer.Image = global::PowerPointAddIn1.Properties.Resources.delete;
            this.buttonRemoveAnswer.Label = "  Hide  ";
            this.buttonRemoveAnswer.Name = "buttonRemoveAnswer";
            this.buttonRemoveAnswer.ShowImage = true;
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.LARS);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.LARS.ResumeLayout(false);
            this.LARS.PerformLayout();
            this.groupCreateNewSurvey.ResumeLayout(false);
            this.groupCreateNewSurvey.PerformLayout();
            this.groupConnect.ResumeLayout(false);
            this.groupConnect.PerformLayout();
            this.groupSelectSurvey.ResumeLayout(false);
            this.groupSelectSurvey.PerformLayout();
            this.startSurveyGroup.ResumeLayout(false);
            this.startSurveyGroup.PerformLayout();
            this.addQuestionGroup.ResumeLayout(false);
            this.addQuestionGroup.PerformLayout();
            this.answerGroup.ResumeLayout(false);
            this.answerGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab LARS;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupCreateNewSurvey;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateNewSurvey;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupConnect;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton connectBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSelectSurvey;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown lectureDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown chapterDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown surveyDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup startSurveyGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton startSurveyButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup addQuestionGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddQuestion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRemoveQuestion;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup answerGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddAnswer;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRemoveAnswer;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon Ribbon
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
