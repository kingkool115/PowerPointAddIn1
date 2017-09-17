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
            this.groupConnect = this.Factory.CreateRibbonGroup();
            this.groupSelectSurvey = this.Factory.CreateRibbonGroup();
            this.select_lecture_group = this.Factory.CreateRibbonGroup();
            this.startSurveyGroup = this.Factory.CreateRibbonGroup();
            this.addQuestionGroup = this.Factory.CreateRibbonGroup();
            this.questions_label = this.Factory.CreateRibbonLabel();
            this.questions_counter = this.Factory.CreateRibbonLabel();
            this.answerGroup = this.Factory.CreateRibbonGroup();
            this.evaluations_label = this.Factory.CreateRibbonLabel();
            this.evaluation_counter = this.Factory.CreateRibbonLabel();
            this.checkGroup = this.Factory.CreateRibbonGroup();
            this.btnCreateNewSurvey = this.Factory.CreateRibbonButton();
            this.connectBtn = this.Factory.CreateRibbonButton();
            this.refreshButton = this.Factory.CreateRibbonButton();
            this.select_lecture_button = this.Factory.CreateRibbonButton();
            this.startSurveyButton = this.Factory.CreateRibbonButton();
            this.button_start_pres_from_slide = this.Factory.CreateRibbonButton();
            this.buttonAddQuestion = this.Factory.CreateRibbonButton();
            this.buttonAddAnswer = this.Factory.CreateRibbonButton();
            this.check_button = this.Factory.CreateRibbonButton();
            this.LARS.SuspendLayout();
            this.groupCreateNewSurvey.SuspendLayout();
            this.groupConnect.SuspendLayout();
            this.groupSelectSurvey.SuspendLayout();
            this.select_lecture_group.SuspendLayout();
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
            this.LARS.Groups.Add(this.select_lecture_group);
            this.LARS.Groups.Add(this.startSurveyGroup);
            this.LARS.Groups.Add(this.addQuestionGroup);
            this.LARS.Groups.Add(this.answerGroup);
            this.LARS.Groups.Add(this.checkGroup);
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
            this.groupConnect.Label = "Not Connected  ";
            this.groupConnect.Name = "groupConnect";
            // 
            // groupSelectSurvey
            // 
            this.groupSelectSurvey.Items.Add(this.refreshButton);
            this.groupSelectSurvey.Label = "Sync Data";
            this.groupSelectSurvey.Name = "groupSelectSurvey";
            // 
            // select_lecture_group
            // 
            this.select_lecture_group.Items.Add(this.select_lecture_button);
            this.select_lecture_group.Label = "Lecture: ---";
            this.select_lecture_group.Name = "select_lecture_group";
            // 
            // startSurveyGroup
            // 
            this.startSurveyGroup.Items.Add(this.startSurveyButton);
            this.startSurveyGroup.Items.Add(this.button_start_pres_from_slide);
            this.startSurveyGroup.Label = "     Start Survey     ";
            this.startSurveyGroup.Name = "startSurveyGroup";
            // 
            // addQuestionGroup
            // 
            this.addQuestionGroup.Items.Add(this.buttonAddQuestion);
            this.addQuestionGroup.Items.Add(this.questions_label);
            this.addQuestionGroup.Items.Add(this.questions_counter);
            this.addQuestionGroup.Label = "  Question  ";
            this.addQuestionGroup.Name = "addQuestionGroup";
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
            // checkGroup
            // 
            this.checkGroup.Items.Add(this.check_button);
            this.checkGroup.Label = "Check";
            this.checkGroup.Name = "checkGroup";
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
            this.connectBtn.Label = "     ";
            this.connectBtn.Name = "connectBtn";
            this.connectBtn.ShowImage = true;
            this.connectBtn.Tag = "connect";
            this.connectBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.connectBtn_Click);
            // 
            // refreshButton
            // 
            this.refreshButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.refreshButton.Enabled = false;
            this.refreshButton.Image = global::PowerPointAddIn1.Properties.Resources.refresh_button;
            this.refreshButton.Label = "        ";
            this.refreshButton.Name = "refreshButton";
            this.refreshButton.ShowImage = true;
            this.refreshButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.refreshButton_Click);
            // 
            // select_lecture_button
            // 
            this.select_lecture_button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.select_lecture_button.Image = global::PowerPointAddIn1.Properties.Resources.lecture;
            this.select_lecture_button.Label = "          ";
            this.select_lecture_button.Name = "select_lecture_button";
            this.select_lecture_button.ShowImage = true;
            this.select_lecture_button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.select_lecture_button_Click);
            // 
            // startSurveyButton
            // 
            this.startSurveyButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.startSurveyButton.Image = ((System.Drawing.Image)(resources.GetObject("startSurveyButton.Image")));
            this.startSurveyButton.Label = "Start Presentatino From Beginning";
            this.startSurveyButton.Name = "startSurveyButton";
            this.startSurveyButton.ShowImage = true;
            this.startSurveyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.startSurveyButton_Click);
            // 
            // button_start_pres_from_slide
            // 
            this.button_start_pres_from_slide.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_start_pres_from_slide.Image = global::PowerPointAddIn1.Properties.Resources.start_presentation_from_this;
            this.button_start_pres_from_slide.Label = " Start Presentation From Current Slide";
            this.button_start_pres_from_slide.Name = "button_start_pres_from_slide";
            this.button_start_pres_from_slide.ShowImage = true;
            this.button_start_pres_from_slide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_start_pres_from_slide_Click);
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
            this.select_lecture_group.ResumeLayout(false);
            this.select_lecture_group.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup startSurveyGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton startSurveyButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup addQuestionGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddQuestion;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup answerGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddAnswer;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup checkGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton check_button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_start_pres_from_slide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton refreshButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel questions_counter;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel questions_label;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel evaluation_counter;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel evaluations_label;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton select_lecture_button;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup select_lecture_group;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon Ribbon
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
