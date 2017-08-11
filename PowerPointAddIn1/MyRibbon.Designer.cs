using PowerPointAddIn1.utils;
using System.Runtime.InteropServices;
using System;

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
            pptNavigator = new PowerPointNavigator();
            myRestHelper = new RestHelperLARS();
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
            this.btnCreateNewSurvey = this.Factory.CreateRibbonButton();
            this.groupConnect = this.Factory.CreateRibbonGroup();
            this.connectBtn = this.Factory.CreateRibbonButton();
            this.groupSelectSurvey = this.Factory.CreateRibbonGroup();
            this.lectureDropDown = this.Factory.CreateRibbonDropDown();
            this.chapterDropDown = this.Factory.CreateRibbonDropDown();
            this.surveyDropDown = this.Factory.CreateRibbonDropDown();
            this.refreshButton = this.Factory.CreateRibbonButton();
            this.startSurveyGroup = this.Factory.CreateRibbonGroup();
            this.startSurveyButton = this.Factory.CreateRibbonButton();
            this.button_start_pres_from_slide = this.Factory.CreateRibbonButton();
            this.addQuestionGroup = this.Factory.CreateRibbonGroup();
            this.buttonAddQuestion = this.Factory.CreateRibbonButton();
            this.questions_label = this.Factory.CreateRibbonLabel();
            this.questions_counter = this.Factory.CreateRibbonLabel();
            this.answerGroup = this.Factory.CreateRibbonGroup();
            this.evaluations_label = this.Factory.CreateRibbonLabel();
            this.evaluation_counter = this.Factory.CreateRibbonLabel();
            this.buttonAddAnswer = this.Factory.CreateRibbonButton();
            this.checkGroup = this.Factory.CreateRibbonGroup();
            this.check_button = this.Factory.CreateRibbonButton();
            this.group_save = this.Factory.CreateRibbonGroup();
            this.LARS.SuspendLayout();
            this.groupCreateNewSurvey.SuspendLayout();
            this.groupConnect.SuspendLayout();
            this.groupSelectSurvey.SuspendLayout();
            this.startSurveyGroup.SuspendLayout();
            this.addQuestionGroup.SuspendLayout();
            this.answerGroup.SuspendLayout();
            this.checkGroup.SuspendLayout();
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
            this.LARS.Groups.Add(this.checkGroup);
            this.LARS.Groups.Add(this.group_save);
            this.LARS.Label = "LARS";
            this.LARS.Name = "LARS";
            // 
            // groupCreateNewSurvey
            // 
            this.groupCreateNewSurvey.Items.Add(this.btnCreateNewSurvey);
            this.groupCreateNewSurvey.Label = "  New Survey  ";
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
            this.groupConnect.Label = "  Not Connected  ";
            this.groupConnect.Name = "groupConnect";
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
            // groupSelectSurvey
            // 
            this.groupSelectSurvey.Items.Add(this.lectureDropDown);
            this.groupSelectSurvey.Items.Add(this.chapterDropDown);
            this.groupSelectSurvey.Items.Add(this.surveyDropDown);
            this.groupSelectSurvey.Items.Add(this.refreshButton);
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
            // refreshButton
            // 
            this.refreshButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.refreshButton.Enabled = false;
            this.refreshButton.Image = global::PowerPointAddIn1.Properties.Resources.refresh;
            this.refreshButton.Label = " ";
            this.refreshButton.Name = "refreshButton";
            this.refreshButton.ShowImage = true;
            this.refreshButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.refreshButton_Click);
            // 
            // startSurveyGroup
            // 
            this.startSurveyGroup.Items.Add(this.startSurveyButton);
            this.startSurveyGroup.Items.Add(this.button_start_pres_from_slide);
            this.startSurveyGroup.Label = "  Start Survey  ";
            this.startSurveyGroup.Name = "startSurveyGroup";
            // 
            // startSurveyButton
            // 
            this.startSurveyButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.startSurveyButton.Image = ((System.Drawing.Image)(resources.GetObject("startSurveyButton.Image")));
            this.startSurveyButton.Label = "Start presentatino from Beginning";
            this.startSurveyButton.Name = "startSurveyButton";
            this.startSurveyButton.ShowImage = true;
            this.startSurveyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.startSurveyButton_Click);
            // 
            // button_start_pres_from_slide
            // 
            this.button_start_pres_from_slide.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_start_pres_from_slide.Image = global::PowerPointAddIn1.Properties.Resources.start_presentation_from_current_slide;
            this.button_start_pres_from_slide.Label = " Start presentation from current slide";
            this.button_start_pres_from_slide.Name = "button_start_pres_from_slide";
            this.button_start_pres_from_slide.ShowImage = true;
            this.button_start_pres_from_slide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_start_pres_from_slide_Click);
            // 
            // addQuestionGroup
            // 
            this.addQuestionGroup.Items.Add(this.buttonAddQuestion);
            this.addQuestionGroup.Items.Add(this.questions_label);
            this.addQuestionGroup.Items.Add(this.questions_counter);
            this.addQuestionGroup.Label = "  Question  ";
            this.addQuestionGroup.Name = "addQuestionGroup";
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
            // questions_label
            // 
            this.questions_label.Label = "    Questions    ";
            this.questions_label.Name = "questions_label";
            // 
            // questions_counter
            // 
            this.questions_counter.Label = "           0";
            this.questions_counter.Name = "questions_counter";
            this.questions_counter.ScreenTip = "   ";
            // 
            // answerGroup
            // 
            this.answerGroup.Items.Add(this.evaluations_label);
            this.answerGroup.Items.Add(this.evaluation_counter);
            this.answerGroup.Items.Add(this.buttonAddAnswer);
            this.answerGroup.Label = "  Evaluation  ";
            this.answerGroup.Name = "answerGroup";
            // 
            // evaluations_label
            // 
            this.evaluations_label.Label = "    Evaluations    ";
            this.evaluations_label.Name = "evaluations_label";
            // 
            // evaluation_counter
            // 
            this.evaluation_counter.Label = "             0";
            this.evaluation_counter.Name = "evaluation_counter";
            // 
            // buttonAddAnswer
            // 
            this.buttonAddAnswer.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAddAnswer.Enabled = false;
            this.buttonAddAnswer.Image = global::PowerPointAddIn1.Properties.Resources.add_answer;
            this.buttonAddAnswer.Label = "  Show  ";
            this.buttonAddAnswer.Name = "buttonAddAnswer";
            this.buttonAddAnswer.ShowImage = true;
            this.buttonAddAnswer.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonEvaluateQuestion_Click);
            // 
            // checkGroup
            // 
            this.checkGroup.Items.Add(this.check_button);
            this.checkGroup.Label = "Check";
            this.checkGroup.Name = "checkGroup";
            // 
            // check_button
            // 
            this.check_button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.check_button.Enabled = false;
            this.check_button.Image = global::PowerPointAddIn1.Properties.Resources.check_questions_image;
            this.check_button.Label = "     ";
            this.check_button.Name = "check_button";
            this.check_button.ShowImage = true;
            this.check_button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.check_button_Click);
            // 
            // group_save
            // 
            this.group_save.Label = "Save";
            this.group_save.Name = "group_save";
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
            this.checkGroup.ResumeLayout(false);
            this.checkGroup.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup answerGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddAnswer;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup checkGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton check_button;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_save;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_start_pres_from_slide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton refreshButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel questions_counter;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel questions_label;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel evaluation_counter;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel evaluations_label;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon Ribbon
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
