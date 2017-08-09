using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RestSharp;
using RestSharp.Authenticators;
using Newtonsoft.Json;
using System.Net;
using System.IO;

namespace PowerPointAddIn1.utils
{
    public class RestHelperLARS
    {

        private String REST_API_URL = "http://127.0.0.1:8000/";
        private RestClient client;
        public bool IsAuthenticated { get; set; }
        String userEmail = "presentation_user";

        /*
         * Constructor.
         */
        public RestHelperLARS()
        {
            client = new RestClient(REST_API_URL);
        }

        /*
         * Explicit authentication throug LoginForm.
         */
        public void authenticate(String username, String password)
        {
            client.Authenticator = new HttpBasicAuthenticator(username, password);
            var request = new RestRequest("/create_new_survey", Method.GET);
            request.AddHeader("accept", "application/json");

            IRestResponse response = client.Execute(request);
            if (response.StatusCode == HttpStatusCode.OK)
            {
                IsAuthenticated = true;
                userEmail = username;
            }
        }

        /*
         * Explicit logout.
         */
        public void logout()
        {
            var request = new RestRequest("logout", Method.GET);
            // execute the request
            client.Execute(request);
            client.Authenticator = null;
            IsAuthenticated = false;
            userEmail = "presentatation_user";
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

        /*
         * Make a new presentation session entry into DB if WebService
         */
        public void startPresentationSession(String sessionId, int lectureId, int chapterId)
        {
            var request = new RestRequest("/start_presentation_session", Method.GET);
            request.AddParameter("session_id", sessionId, ParameterType.GetOrPost);
            request.AddParameter("lecture_id", lectureId, ParameterType.GetOrPost);
            request.AddParameter("chapter_id", chapterId, ParameterType.GetOrPost);
            request.AddParameter("user_email", userEmail, ParameterType.GetOrPost);
            request.AddHeader("Content-Type", "application/json");

            // execute the request
            request.RequestFormat = DataFormat.Json;
            IRestResponse response = client.Execute(request);
            return;
        }

        /*
         * Get evaluation of given questions for a certain sessionId
         */
        public List<Evaluation> EvaluateAnswers(List<String> questionsIds, String sessionId)
        {
            // concatenate question ids (workaround)
            String concatQuestionIds = "";
            foreach (var id in questionsIds)
            {
                concatQuestionIds += id + ",";
            }

            // create request
            var request = new RestRequest("/evaluate_answers", Method.GET);
            request.AddParameter("session_id", sessionId, ParameterType.GetOrPost);
            request.AddParameter("question_ids", concatQuestionIds, ParameterType.GetOrPost);
            request.AddHeader("Content-Type", "application/json");

            // execute the request
            request.RequestFormat = DataFormat.Json;
            IRestResponse response = client.Execute(request);
            var evaluationList = JsonConvert.DeserializeObject<List<Evaluation>>(response.Content);
            return evaluationList;
        }

        /*
         * Push a question to the registered devices. 
         */
        public void pushQuestion(int questionId, String sessionId, String lectureId, String userEmail)
        {
            // create request
            var request = new RestRequest("/push_question", Method.GET);
        }
    }
}
