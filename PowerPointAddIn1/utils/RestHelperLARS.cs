using System;
using System.Collections.Generic;
using RestSharp;
using RestSharp.Authenticators;
using Newtonsoft.Json;
using System.Net;
using System.IO;
using System.Drawing;

namespace PowerPointAddIn1.utils
{
    public class RestHelperLARS
    {

        //TODO: change URL
        private String REST_API_URL = "http://127.0.0.1:8000/";
        private RestClient client;
        public bool IsAuthenticated { get; set; }
        public String userEmail = "presentation_user";

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
         * Get all lectures of all users. 
         */
        public IRestResponse getAllAvailableLectures()
        {
            var request = new RestRequest("api/all_lectures", Method.GET);
            request.AddHeader("Content-Type", "application/json");
            request.AddParameter("user_email", userEmail, ParameterType.GetOrPost);
            request.RequestFormat = DataFormat.Json;
            // execute the request
            return client.Execute(request);
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
         * Get all chapters of a certain lecture.
         */
        public List<Chapter> GetChaptersOfLectureAsGuest(String lectureId)
        {
            // create request
            var request = new RestRequest("/api/lecture/{lecture_id}/all_chapters", Method.GET);
            request.AddHeader("Content-Type", "application/json");
            request.AddUrlSegment("lecture_id", "" + lectureId); // replaces matching token in request.Resource
            request.AddParameter("user_email", userEmail, ParameterType.GetOrPost);
            request.RequestFormat = DataFormat.Json;

            // execute the request
            IRestResponse response = client.Execute(request);
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
        public IRestResponse startPresentationSession(String sessionId, int? lectureId)
        {
            var request = new RestRequest("/api/start_presentation_session", Method.GET);
            request.AddParameter("session_id", sessionId, ParameterType.GetOrPost);
            request.AddParameter("lecture_id", lectureId, ParameterType.GetOrPost);
            request.AddParameter("user_email", userEmail, ParameterType.GetOrPost);
            request.AddHeader("Content-Type", "application/json");

            // execute the request
            request.RequestFormat = DataFormat.Json;
            IRestResponse response = client.Execute(request);
            return response;
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
            var request = new RestRequest("/api/evaluate_answers", Method.GET);
            request.AddParameter("session_id", sessionId, ParameterType.GetOrPost);
            request.AddParameter("question_ids", concatQuestionIds, ParameterType.GetOrPost);
            request.AddHeader("Content-Type", "application/json");

            // execute the request
            request.RequestFormat = DataFormat.Json;
            IRestResponse response = client.Execute(request);
            JsonSerializerSettings settings = new JsonSerializerSettings();
            try
            {
                var evaluationList = JsonConvert.DeserializeObject<List<Evaluation>>(response.Content, settings);
                return evaluationList;
            }
            catch (JsonSerializationException ex) {
                Console.WriteLine("Could not deserialize " + response.Content + " to Evaluation Object. Maybe because answers field is empty.");
                return null;
            }
        }

        /*
         * Push a question to the registered devices. 
         */
        public IRestResponse pushQuestion(String questionId, String lectureId, String sessionId, String userEmail)
        {
            // create request
            var request = new RestRequest("/api/push_question", Method.GET);
            request.AddParameter("question_id", questionId, ParameterType.GetOrPost);
            request.AddParameter("lecture_id", lectureId, ParameterType.GetOrPost);
            request.AddParameter("session_id", sessionId, ParameterType.GetOrPost);
            request.AddParameter("user_email", userEmail, ParameterType.GetOrPost);
            request.AddHeader("Content-Type", "application/json");

            // execute the request
            request.RequestFormat = DataFormat.Json;
            IRestResponse response = client.Execute(request);
            return response;
        }

        /*
         * Download an image which of an question.
         */
        public String downloadQuestionImage(String imagePath, String imageUrl) {

            // download image to bitmap
            WebClient client = new WebClient();
            Stream stream = client.OpenRead(imageUrl);
            Bitmap bitmap; bitmap = new Bitmap(stream);

            // get extension of image
            string url = imageUrl;
            string ext = Path.GetExtension(url);

            // save image
            String filename = imagePath + ext;
            if (bitmap != null)
                bitmap.Save(filename);

            stream.Flush();
            stream.Close();
            client.Dispose();
            return filename;
        }

        /*
         * Gets the answers of the students to a question.
         */
        public Dictionary<String, int> GetAnswersForQuestion(string questionId, string sessionId)
        {
            // create request
            var request = new RestRequest("/api/get_answers_of_one_question/{question_id}/{session_id}", Method.GET);
            request.AddHeader("accept", "application/json");
            request.AddUrlSegment("question_id", "" + questionId);
            request.AddUrlSegment("session_id", "" + sessionId);    // replaces matching token in request.Resource

            // execute the request
            IRestResponse response = client.Execute(request);
            var evaluation = JsonConvert.DeserializeObject<Evaluation>(response.Content);
            return evaluation.Answers;
        }
    }
}
