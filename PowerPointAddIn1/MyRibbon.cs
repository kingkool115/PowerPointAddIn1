using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using RestSharp;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Tools.Ribbon;
using RestSharp.Authenticators;
using Newtonsoft.Json;
using PowerPointAddIn1.utils;

namespace PowerPointAddIn1
{
    public partial class MyRibbon
    {
        public RestHelperLARS myRestHelper;

        public String REST_API_URL = "http://127.0.0.1:8000/";
        public String username;
        public String password;


        public void initRestHelper(RestHelperLARS restHelper)
        {
            myRestHelper = restHelper;
        }

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {            
            // TODO: checken ob User eingeloggt ist. Wenn nicht, dann wird selectSurvey-group ausgegraut.
        }

        private void btnCreateNewSurvey_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("http://127.0.0.1:8000/create_new_survey");
        }

        /*
         * Enable/disble Ribbon buttons after login/logout.
         */
        public void enableRibbons(Boolean enable)
        {
            lectureDropDown.Enabled = enable;
            chapterDropDown.Enabled = enable;
            surveyDropDown.Enabled = enable;
            buttonAddQuestion.Enabled = enable;
            buttonRemoveQuestion.Enabled = enable;
            buttonAddAnswer.Enabled = enable;
            buttonRemoveAnswer.Enabled = enable;
           
        }

        /*
         * Logout from LARS.
         */
        private void doLogout()
        {
            var client = new RestClient(REST_API_URL);
            var request = new RestRequest("logout", Method.GET);
            // execute the request
            client.Execute(request);
            myRestHelper = null;
        }

        /*
         * Handles click on Connect Button. 
         */
        private void connectBtn_Click(object sender, RibbonControlEventArgs e)
        {
            // click to connect
            if (connectBtn.Tag.Equals("connect"))
            {
                LoginForm loginForm = new LoginForm();
                loginForm.Show();
            }
            // click to disconnect
            else
            {
                connectBtn.Image = PowerPointAddIn1.Properties.Resources.connect;
                groupConnect.Label = "Not Connected";
                doLogout();
                enableRibbons(false);   // disable ribbons
                connectBtn.Tag = "connect";
            }
        }

        public void lectureDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            chapterDropDown.Items.Clear();
            surveyDropDown.Items.Clear();
            
            String selectedLectureId = (String) lectureDropDown.SelectedItem.Tag;
             
            var response = myRestHelper.getChaptersOfLecture(selectedLectureId);

            var chapterList = JsonConvert.DeserializeObject<List<Chapter>>(response.Content);

            // fill lecture combobox
            foreach (var chapter in chapterList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = chapter.Name;
                item.Tag = chapter.ID;
                chapterDropDown.Items.Add(item);
            }
        }

        public void chapterDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            surveyDropDown.Items.Clear();

            String selectedLectureId = (String) lectureDropDown.SelectedItem.Tag;
            String selectedChapterId = (String) chapterDropDown.SelectedItem.Tag;

            // execute the request
            IRestResponse response = myRestHelper.getSurveysOfChapter(selectedLectureId, selectedChapterId);
            
            // TODO: if content there are no surveys for that chapter ...
            var surveysList = JsonConvert.DeserializeObject<List<Survey>>(response.Content);

            // fill lecture combobox
            foreach (var survey in surveysList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = survey.Name;
                item.Tag = survey.ID;
                surveyDropDown.Items.Add(item);
            }
        }

        private void buttonAddQuestion_Click(object sender, RibbonControlEventArgs e)
        {
            SelectQuestionsForm selectQuestions = new SelectQuestionsForm();
            selectQuestions.Show();
        }
    }
}
