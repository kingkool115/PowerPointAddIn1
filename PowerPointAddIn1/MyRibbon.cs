using System;
using System.Collections.Generic;
using RestSharp;
using Microsoft.Office.Tools.Ribbon;
using PowerPointAddIn1.utils;
using System.Windows.Forms;
using Newtonsoft.Json;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    public partial class MyRibbon
    {
        // connects to REST Service where questions are stored
        public RestHelperLARS myRestHelper;

        // observes the slide navigation
        public PowerPointNavigator pptNavigator;
        
        // SessionController needed when a presentation is started.
        public SessionController SessionController { get; set; }

        // the form to add/remove questions to a slide
        public SelectQuestionsForm selectQuestionsForm;

        // the form to evaluate questions on a slide
        public EvaluateQuestionsForm evaluateQuestionsForm;
        
        public Lecture LectureForThisPresentation { get; set; }

        // needed for comboboxes in ribbon
        public List<Lecture> myLectures;
        public Lecture currentLecture;
        public Chapter currentChapter;

        // setting on true when running a survey session
        public Boolean isSessionRunning { get; set; }

        bool isUserLoggedIn;

        /*
         * Update the counter in the ribbon.
         */
        public void updateRibbonQuestionEvaluationCounter(int slideId)
        {
            var customSlide = pptNavigator.getCustomSlideById(slideId);
            if (customSlide != null)
            {
                questions_counter.Label = "             " + customSlide.PushQuestionList.Count;
                evaluation_counter.Label = "             " + customSlide.EvaluationList.Count;
            }
            else
            {
                questions_counter.Label = "              0";
                evaluation_counter.Label = "              0";
            }
        }
        
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            // TODO: checken ob User eingeloggt ist. Wenn nicht, dann wird selectSurvey-group ausgegraut.
        }

        /*
         * Redirects to create new survey website in LARS.
         */
        private void btnCreateNewSurvey_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("http://127.0.0.1:8000/create_new_survey");
        }

        /*
         * Enable/disble Ribbon buttons after login/logout.
         */
        public void enableRibbons(Boolean enable)
        {
            buttonAddQuestion.Enabled = enable;
            buttonAddAnswer.Enabled = enable;
            check_button.Enabled = enable;
            refreshButton.Enabled = enable;
        }
                

        /*
         * This method is called after successful login.
         * Inits myRestHelper, fills lecturesList, enables RibbonButtons and fills lecture dropdown list in Ribbon.
         */
        public void afterSuccessfulLogin(List<Lecture> lectureList)
        {
            // enable ribbons
            enableRibbons(true);
            
            initLectures(lectureList);

            // change Connect-Button to Disconnect
            connectBtn.Image = Properties.Resources.disconnect;
            connectBtn.Tag = "disconnect";
            groupConnect.Label = "Connected";
            isUserLoggedIn = true;
        }

        /*
         * Is called after successful login or after pressing refresh button.
         */
        public void initLectures(List<Lecture> lectureList)
        {
            this.myLectures = lectureList;

            // init lectures, chapters, surveys and questions
            foreach (var lecture in lectureList)
            {
                var chapterList = myRestHelper.GetChaptersOfLecture(lecture.ID);
                lecture.setChapters(chapterList);

                foreach (var chapter in chapterList)
                {
                    var surveyList = myRestHelper.GetSurveysOfChapter(lecture.ID, chapter.ID);
                    chapter.setSurveys(surveyList);

                    foreach (var survey in surveyList)
                    {
                        var questionList = myRestHelper.GetQuestionsOfSurvey(lecture.ID, chapter.ID, survey.ID);
                        survey.setQuestions(questionList);

                        foreach (var question in survey.getQuestions())
                        {
                            question.setLectureChapterSurvey(lecture, chapter, survey);
                        }
                    }
                }
            }
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
                connectBtn.Image = Properties.Resources.connect;
                groupConnect.Label = "Not Connected";
                myRestHelper.logout();
                enableRibbons(false);   // disable ribbons
                connectBtn.Tag = "connect";
                isUserLoggedIn = false;
            }
        }
        
        /*
         * Get a lecture from myLectures by id.
         */
        public Lecture getLectureById(String lectureId)
        {
            foreach (var lecture in myLectures)
            {
                if (lecture.ID.Equals(lectureId))
                {
                    return lecture;
                }
            }
            return null;
        }
        
        /*
         * Get a lecture from myLectures by id.
         */
        private Chapter getChapterById(String chapterId)
        {
            foreach (var chapter in currentLecture.getChapters())
            {
                if (chapter.ID.Equals(chapterId))
                {
                    return chapter;
                }
            }
            return null;
        }
        
        /*
         * This function is called AddQuestion button is clicked.
         * It opens SelectQuestionsForm.
         */
        private void buttonAddQuestion_Click(object sender, RibbonControlEventArgs e)
        {
            selectQuestionsForm = new SelectQuestionsForm();
            selectQuestionsForm.Show();
            selectQuestionsForm.updateQuestionsPerSlideListView();
        }

        /*
         * This function is called AddQuestion button is clicked.
         * It opens SelectQuestionsForm.
         */
        private void buttonEvaluateQuestion_Click(object sender, RibbonControlEventArgs e)
        {
            evaluateQuestionsForm = new EvaluateQuestionsForm();
            evaluateQuestionsForm.Show();
            //selectQuestionsForm.updateQuestionsPerSlideListView();
        }

        /*
         * Check button clicked.
         * Check if pushed questions will ever be evaluated, if they are pushed/evaluated in the given order.
         */
        private void check_button_Click(object sender, RibbonControlEventArgs e)
        {
            List<String> errorMessages = pptNavigator.checkQuestionsPushEvaluationOrder(true);

            // if there is any errorMessage than display popup
            if (errorMessages.Count > 0)
            {
                String errMessage = "";
                foreach (var message in errorMessages)
                {
                    errMessage = errMessage + message + "\n\n";
                }
                MessageBox.Show(errMessage, "You have pushed some questions which will never be evaluated.");
            }
            else
            {
                MessageBox.Show("Your push/evaluation order of your questions is fine !!!");
            }
        }

        /*
         * Start presentation from first slide with preconfigured custom slides.
         */
        private void startSurveyButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (isSessionRunning)
            {
                DialogResult dialogResult = MessageBox.Show("All evaluation slides will be removed when you finish this session.\nAre you sure?",
                    "Finish presentation session.", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    //enable ribbon buttons
                    if (isUserLoggedIn)
                    {
                        buttonAddQuestion.Enabled = true;
                        buttonAddAnswer.Enabled = true;
                        check_button.Enabled = true;
                        refreshButton.Enabled = true;
                    }
                    button_start_pres_from_slide.Enabled = true;
                    connectBtn.Enabled = true;
                    select_lecture_button.Enabled = true;

                    startSurveyButton.Image = Properties.Resources.play_sign;
                    isSessionRunning = false;
                    SessionController.removeEventHandlers();
                    foreach (var evaluationSlide in SessionController.EvaluationSlides.Keys)
                    {
                        pptNavigator.slides[evaluationSlide.SlideIndex].Delete();
                    }
                }
            }
            else
            {
                openSessionForm(true);
            }
        }

        /*
         * Start presentation from currently selected slide with preconfigured custom slides.
         */
        private void button_start_pres_from_slide_Click(object sender, RibbonControlEventArgs e)
        {
            openSessionForm(false);
        }

        /*
         * Open StartSessionForm.
         */
        private void openSessionForm(bool fromBeginning)
        {
            var response = myRestHelper.getAllAvailableLectures();
            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                DialogResult dialogResult = MessageBox.Show("Could not connect to the ARS server.",
                "Connection Failed", MessageBoxButtons.OK);
                return;
            }
            StartSessionForm sessionForm = new StartSessionForm(this, fromBeginning, isUserLoggedIn, response);
            sessionForm.Show();
        }

        /*
         * Start a new session. 
         */
        public void startNewSession(String selectedLectureId, bool fromBeginning, int timeSpentOnSlideBeforePushing)
        {
            SessionController = new SessionController(this, pptNavigator.pptApplication);
            bool sessionStartedSuccessfully = SessionController.startPresentation(fromBeginning, pptNavigator.SlideIndex,
                                                pptNavigator.presentation, pptNavigator.slides,
                                                selectedLectureId, timeSpentOnSlideBeforePushing);
            if (sessionStartedSuccessfully)
            {
                startSurveyButton.Image = Properties.Resources.stop;
                isSessionRunning = true;

                // disable ribbon buttons during session
                if (isUserLoggedIn)
                {
                    buttonAddQuestion.Enabled = false;
                    buttonAddAnswer.Enabled = false;
                    check_button.Enabled = false;
                    refreshButton.Enabled = false;
                }
                button_start_pres_from_slide.Enabled = false;
                connectBtn.Enabled = false;
                select_lecture_button.Enabled = false;
            }
        }
        
        private void refreshButton_Click(object sender, RibbonControlEventArgs e)
        {
            // execute the request
            IRestResponse response = myRestHelper.getAllLectures();

            // if login is successfull
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                var content = response.Content;
                var lectureList = JsonConvert.DeserializeObject<List<Lecture>>(content);
                initLectures(lectureList);
            }
        }

        /*
         * Show SelectLectureForm.
         */
        private void select_lecture_button_Click(object sender, RibbonControlEventArgs e)
        {
            SelectLectureForm selectLectureForm = new SelectLectureForm(this);
            selectLectureForm.Show();
        }
    }
}
