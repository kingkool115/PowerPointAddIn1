namespace PowerPointAddIn1
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon));
            this.LARS = this.Factory.CreateRibbonTab();
            this.groupCreateNewSurvey = this.Factory.CreateRibbonGroup();
            this.btnCreateNewSurvey = this.Factory.CreateRibbonButton();
            this.groupConnect = this.Factory.CreateRibbonGroup();
            this.connectBtn = this.Factory.CreateRibbonButton();
            this.groupSelectSurvey = this.Factory.CreateRibbonGroup();
            this.lectureDropDown = this.Factory.CreateRibbonDropDown();
            this.chapterDropDown = this.Factory.CreateRibbonDropDown();
            this.surveyDropDown = this.Factory.CreateRibbonDropDown();
            this.startSurveyGroup = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.addQuestionGroup = this.Factory.CreateRibbonGroup();
            this.buttonAddQuestion = this.Factory.CreateRibbonButton();
            this.LARS.SuspendLayout();
            this.groupCreateNewSurvey.SuspendLayout();
            this.groupConnect.SuspendLayout();
            this.groupSelectSurvey.SuspendLayout();
            this.startSurveyGroup.SuspendLayout();
            this.addQuestionGroup.SuspendLayout();
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
            this.LARS.Label = "LARS";
            this.LARS.Name = "LARS";
            // 
            // groupCreateNewSurvey
            // 
            this.groupCreateNewSurvey.Items.Add(this.btnCreateNewSurvey);
            this.groupCreateNewSurvey.Label = "New Survey";
            this.groupCreateNewSurvey.Name = "groupCreateNewSurvey";
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
            // groupConnect
            // 
            this.groupConnect.Items.Add(this.connectBtn);
            this.groupConnect.Label = "Not Connected";
            this.groupConnect.Name = "groupConnect";
            // 
            // connectBtn
            // 
            this.connectBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.connectBtn.Image = ((System.Drawing.Image)(resources.GetObject("connectBtn.Image")));
            this.connectBtn.Label = "Connect";
            this.connectBtn.Name = "connectBtn";
            this.connectBtn.ShowImage = true;
            this.connectBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.connectBtn_Click);
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
            this.lectureDropDown.Label = "Lecture: ";
            this.lectureDropDown.Name = "lectureDropDown";
            this.lectureDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.lectureDropDown_SelectionChanged);
            // 
            // chapterDropDown
            // 
            this.chapterDropDown.Label = "Chapter: ";
            this.chapterDropDown.Name = "chapterDropDown";
            this.chapterDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chapterDropDown_SelectionChanged);
            // 
            // surveyDropDown
            // 
            this.surveyDropDown.Label = "Survey:  ";
            this.surveyDropDown.Name = "surveyDropDown";
            // 
            // startSurveyGroup
            // 
            this.startSurveyGroup.Items.Add(this.button1);
            this.startSurveyGroup.Label = "Start Survey";
            this.startSurveyGroup.Name = "startSurveyGroup";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::PowerPointAddIn1.Properties.Resources.play_sign;
            this.button1.Label = " ";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // addQuestionGroup
            // 
            this.addQuestionGroup.Items.Add(this.buttonAddQuestion);
            this.addQuestionGroup.Label = "Add Question";
            this.addQuestionGroup.Name = "addQuestionGroup";
            // 
            // buttonAddQuestion
            // 
            this.buttonAddQuestion.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAddQuestion.Image = ((System.Drawing.Image)(resources.GetObject("buttonAddQuestion.Image")));
            this.buttonAddQuestion.Label = " ";
            this.buttonAddQuestion.Name = "buttonAddQuestion";
            this.buttonAddQuestion.ShowImage = true;
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup addQuestionGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddQuestion;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
