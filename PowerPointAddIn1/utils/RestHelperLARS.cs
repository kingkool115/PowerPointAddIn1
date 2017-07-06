using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RestSharp;
using RestSharp.Authenticators;
using Newtonsoft.Json;

namespace PowerPointAddIn1.utils
{
    public class RestHelperLARS
    {

        private String REST_API_URL = "http://127.0.0.1:8000/";
        private RestClient client;

        public RestHelperLARS(String username, String password)
        {
            client = new RestClient(REST_API_URL);
            client.Authenticator = new HttpBasicAuthenticator(username, password);
        }

        /*
         * Get all lectures of the user. 
         */
        public IRestResponse getAllLectures()
        {
            var request = new RestRequest("lectures", Method.GET);
            request.AddHeader("accept", "application/json");

            // execute the request
            return client.Execute(request);
        }

        /*
         * Get all chapters of a certain lecture.
         */
        public List<Chapter> GetChaptersOfLecture(String lectureId)
        {
            // create request
            var request = new RestRequest("/lecture/{lecture_id}/chapters", Method.GET);
            request.AddHeader("accept", "application/json");
            request.AddUrlSegment("lecture_id", "" + lectureId); // replaces matching token in request.Resource

            // execute the request
            IRestResponse response =  client.Execute(request);
            var chapterList = JsonConvert.DeserializeObject<List<Chapter>>(response.Content);
            return chapterList;
        }

        /*
         * Get all surveys of a certain chapter.
         */
        public List<Survey> GetSurveysOfChapter(String lectureId, String chapterId)
        {
            // create request
            var request = new RestRequest("/lecture/{lecture_id}/chapter/{chapter_id}/surveys", Method.GET);
            request.AddHeader("accept", "application/json");
            request.AddUrlSegment("lecture_id", "" + lectureId); // replaces matching token in request.Resource
            request.AddUrlSegment("chapter_id", "" + chapterId);

            // execute the request
            IRestResponse response = client.Execute(request);
            var surveyList = JsonConvert.DeserializeObject<List<Survey>>(response.Content);
            return surveyList;
        }

        /*
         * Get all questions of a certain survey.
         */
        public List<Question> GetQuestionsOfSurvey(String lectureId, String chapterId,String surveyId)
        {
            // create request
            var request = new RestRequest("/lecture/{lecture_id}/chapter/{chapter_id}/survey/{survey_id}", Method.GET);
            request.AddHeader("accept", "application/json");
            request.AddUrlSegment("lecture_id", "" + lectureId); // replaces matching token in request.Resource
            request.AddUrlSegment("chapter_id", "" + chapterId);
            request.AddUrlSegment("survey_id", "" + surveyId);

            // execute the request
            IRestResponse response = client.Execute(request);
            var questionList = JsonConvert.DeserializeObject<List<Question>>(response.Content);
            return questionList;
        }

    }
}
