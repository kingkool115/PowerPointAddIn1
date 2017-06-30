using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using RestSharp;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Tools.Ribbon;
using RestSharp.Authenticators;
using Newtonsoft.Json;

namespace PowerPointAddIn1
{

    public partial class Ribbon
    {
        String REST_API_URL = "http://127.0.0.1:8000/" ;
        String USER = "george.handball@web.de";
        String PASSWORD = "123456";

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {            
            // TODO: checken ob User eingeloggt ist. Wenn nicht, dann wird selectSurvey-group ausgegraut.
        }

        private void btnCreateNewSurvey_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("http://127.0.0.1:8000/create_new_survey");
        }

        private void connectBtn_Click(object sender, RibbonControlEventArgs e)
        {
            var btn = sender as RibbonButton;
            var currentImageLabel = btn.Label;

            if (currentImageLabel.Equals("Connect"))
            {
                var client = new RestClient(REST_API_URL);
                client.Authenticator = new HttpBasicAuthenticator(USER, PASSWORD);

                var request = new RestRequest("lectures", Method.GET);
                request.AddHeader("accept", "application/json");

                // execute the request
                IRestResponse response = client.Execute(request);
                var content = response.Content;

                var lectureList = JsonConvert.DeserializeObject<List<Lecture>>(content);

                // fill lecture combobox
                foreach (var lecture in lectureList)
                {
                    RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                    item.Label = lecture.Name;
                    item.Tag = lecture.ID;
                    lectureDropDown.Items.Add(item);
                }

                this.connectBtn.Image = PowerPointAddIn1.Properties.Resources.connected;
                this.connectBtn.Label = "Disconnect";
            }
            else
            {
                this.connectBtn.Image = PowerPointAddIn1.Properties.Resources.not_connected;
                this.connectBtn.Label = "Connect";
            }

            lectureDropDown_SelectionChanged(null, null);
            chapterDropDown_SelectionChanged(null, null);
           // System.Drawing.Bitmap bitmap = PowerPointAddIn1.Properties.Resources.connected;
            //connectBtn.Image = bitmap;
        }

        private void lectureDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            chapterDropDown.Items.Clear();
            surveyDropDown.Items.Clear();

            var selectedLectureId = lectureDropDown.SelectedItem.Tag;

            // init client+authentication
            var client = new RestClient(REST_API_URL);
            client.Authenticator = new HttpBasicAuthenticator(USER, PASSWORD);

            // create request
            var request = new RestRequest("/lecture/{lecture_id}/chapters", Method.GET);
            request.AddHeader("accept", "application/json");
            request.AddUrlSegment("lecture_id", "" + selectedLectureId); // replaces matching token in request.Resource

            // execute the request
            IRestResponse response = client.Execute(request);
            var content = response.Content;

            var chapterList = JsonConvert.DeserializeObject<List<Chapter>>(content);

            // fill lecture combobox
            foreach (var chapter in chapterList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = chapter.Name;
                item.Tag = chapter.ID;
                chapterDropDown.Items.Add(item);
            }
        }

        private void chapterDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            surveyDropDown.Items.Clear();

            var selectedLectureId = lectureDropDown.SelectedItem.Tag;
            var selectedChapterId = chapterDropDown.SelectedItem.Tag;

            // init client+authentication
            var client = new RestClient(REST_API_URL);
            client.Authenticator = new HttpBasicAuthenticator(USER, PASSWORD);

            // create request
            var request = new RestRequest("/lecture/{lecture_id}/chapter/{chapter_id}/surveys", Method.GET);
            request.AddHeader("accept", "application/json");
            request.AddUrlSegment("lecture_id", "" + selectedLectureId); // replaces matching token in request.Resource
            request.AddUrlSegment("chapter_id", "" + selectedChapterId); // replaces matching token in request.Resource

            // execute the request
            IRestResponse response = client.Execute(request);
            var content = response.Content;
            
            // TODO: if content there are no surveys for that chapter ...
            var surveysList = JsonConvert.DeserializeObject<List<Survey>>(content);

            // fill lecture combobox
            foreach (var survey in surveysList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = survey.Name;
                item.Tag = survey.ID;
                surveyDropDown.Items.Add(item);
            }
        }
    }
}
