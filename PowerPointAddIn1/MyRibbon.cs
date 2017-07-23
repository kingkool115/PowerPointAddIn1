using System;
using System.Collections.Generic;
using RestSharp;
using Microsoft.Office.Tools.Ribbon;
using PowerPointAddIn1.utils;
using PPt = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Xml;
using Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public partial class MyRibbon
    {
        // connects to REST Service where questions are stored
        public RestHelperLARS myRestHelper;

        // observes the slide navigation
        public PowerPointNavigator pptNavigator;

        // represents all slides which will push notifications to students
        public List<CustomSlide> questionSlides = new List<CustomSlide>();

        // the form to add/remove questions to a slide
        public SelectQuestionsForm selectQuestionsForm;
        // the form to evaluate questions on a slide
        public EvaluateQuestionsForm evaluateQuestionsForm;

        // needed for comboboxes in ribbon
        public List<Lecture> myLectures;
        public Lecture currentLecture;
        public Chapter currentChapter;

        // REST api stuff
        public String REST_API_URL = "http://127.0.0.1:8000/";
        public String username;
        public String password;
        

        /*
         * Init RestHelper.
         */
        public void initRestHelper(RestHelperLARS restHelper)
        {
            myRestHelper = restHelper;
        }

        /*
         * Check if a CustomSlide for given param slideIndex does already exist in questionSlides.
         */
        public CustomSlide getCustomSlideByIndex(int? slideIndex)
        {

            foreach (var slide in questionSlides)
            {
                if (slide.SlideIndex == slideIndex)
                {
                    return slide;
                }
            }
            return null;
        }

        /*
         * Check if a CustomSlide for given param slideIndex does already exist in questionSlides.
         */
        public CustomSlide getCustomSlideById(int? slideId)
        {
           
            foreach (var slide in questionSlides)
            {
                if (slide.SlideId == slideId)
                {
                    return slide;
                }
            }
            return null;
        }

        /*
         * Is called when Add-Evaluation-Button is clicked in EvaluateQuestionsForm.
         * Provide a slide index to EvaluateSlideIndex-attribute of a question.
         */
        public void addEvaluationToSlide(int slideIdToEvaluate, QuestionObj question)
        {
            if (getCustomSlideById(question.PushSlideId).getQuestion(question) != null)
            {
                getCustomSlideById(question.PushSlideId).getQuestion(question).EvaluateSlideId = slideIdToEvaluate;
            }
        }

        /*
         * Is called when Remove-Evaluation-Button is clicked in EvaluateQuestionsForm.
         * Set EvaluateSlideIndex-attribute of a question to null.
         */
        public void removeEvaluationFromSlide(int slideIdToEvaluate, QuestionObj question)
        {
            // iterate through all custom slides
            for (var x = 1; x < pptNavigator.SlideIndex; x++)
            {
                // find custom slide which have certain question
                if (getCustomSlideByIndex(x).getQuestion(question) != null &&
                    getCustomSlideByIndex(x).getQuestion(question).EvaluateSlideId == slideIdToEvaluate)
                {
                    // remove evaluation on current slide by setting EvaluateSlideIndex to null
                    getCustomSlideByIndex(x).getQuestion(question).EvaluateSlideId = null;
                }
            }
            
        }

        /*
         * Add question to a certain slide.
         */
        public void addQuestionToSlide(int slideId, int slideIndex, QuestionObj question)
        {
            if (getCustomSlideById(slideId) != null)
            {
                // find slide in questionSlides and add question
                getCustomSlideById(slideId).addQuestion(question);
            }
            else
            {
                // create new CustomSlide in questionSlides list
                questionSlides.Add(new CustomSlide(slideId, slideIndex, question));
            }
        }

        /*
         * Removes question from a certain slide.
         */
        public void removeQuestionFromSlide(int slideId, QuestionObj question)
        {
            if (getCustomSlideById(slideId) != null)
            {
                getCustomSlideById(slideId).questionList.Remove(question);
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
            lectureDropDown.Enabled = enable;
            chapterDropDown.Enabled = enable;
            surveyDropDown.Enabled = enable;
            buttonAddQuestion.Enabled = enable;
            buttonAddAnswer.Enabled = enable;
            check_button.Enabled = enable;
            save_button.Enabled = enable;
        }

        /*
         * Init Powerpoint-Navigator.
         */
        public void setPowerpointNavigator(PowerPointNavigator navigator)
        {
            pptNavigator = navigator;
        }

        /*
         * This method is called after successful login.
         * Inits myRestHelper, fills lecturesList, enables RibbonButtons and fills lecture dropdown list in Ribbon.
         **/
        public void doLogin(String username, String password, List<Lecture> lectureList)
        {            
            // init restHelperLARS instance
            initRestHelper(new RestHelperLARS(username, password));

            // enable ribbons
            enableRibbons(true);

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

            // change Connect-Button to Disconnect
            connectBtn.Image = PowerPointAddIn1.Properties.Resources.disconnect;
            connectBtn.Tag = "disconnect";
            groupConnect.Label = "Connected";
            // fill lecture dropdown list
            foreach (var lecture in lectureList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = lecture.Name;
                item.Tag = lecture.ID;
                lectureDropDown.Items.Add(item);
            }

            // fill dropdown lists
            lectureDropDown_SelectionChanged(null, null);
            chapterDropDown_SelectionChanged(null, null);
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
        
        /*
         * Get a lecture from myLectures by id.
         */
        private Lecture getLectureById(String lectureId)
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
         * This method is called whenever a change on lecture dropdown list was made.
         * It fills chapters dropdown list with chapters which belong to the changed lecture.
         */
        public void lectureDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            chapterDropDown.Items.Clear();
            surveyDropDown.Items.Clear();
            
            String selectedLectureId = (String) lectureDropDown.SelectedItem.Tag;
            Lecture lecture = getLectureById(selectedLectureId);
            currentLecture = lecture;

            // fill lecture combobox
            foreach (var chapter in lecture.getChapters())
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = chapter.Name;
                item.Tag = chapter.ID;
                chapterDropDown.Items.Add(item);
            }
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
         * This method is called whenever a change on chapter dropdown list was made.
         * It fills survey dropdown list with chapters which belong to the changed chapter.
         */
        public void chapterDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            surveyDropDown.Items.Clear();
            
            String selectedChapterId = (String) chapterDropDown.SelectedItem.Tag;
            Chapter chapter = getChapterById(selectedChapterId);
            currentChapter = chapter;

            // fill survey combobox
            foreach (var survey in chapter.getSurveys())
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = survey.Name;
                item.Tag = survey.ID;
                surveyDropDown.Items.Add(item);
            }
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
         * Remove a custom slide by its id.
         */
        public void removeCustomSlide(int slideId)
        {
            foreach (var customSlide in questionSlides)
            {
                if (customSlide.SlideId == slideId)
                {
                    questionSlides.Remove(customSlide);
                    break;
                }
            }
        }

        /*
         * Is called whenever slides are added/removed from presentation.
         */
        public void incrementDecrementCustomSlideIndexes(int position, int incDecSlideIndexValue)
        {
            foreach (var customSlide in questionSlides)
            {
                // wenn es sich um einen Slide handelt, dessen index >= ist als das
                // hinzugefügte/gelöschte slide. Denn nur ist der Index des customSlide betroffen.
                if (customSlide.SlideIndex >= position)
                {
                    // new slide index is always higher than 0
                    if (customSlide.SlideIndex + incDecSlideIndexValue > 0)
                    {
                        customSlide.updateSlideIndex(customSlide.SlideIndex + incDecSlideIndexValue);
                    }
                }
            }
        }

        /*
         * Is called whenever slides are dragged & dropped.
         */
        public void updateCustomSlideIndexIfSlideDraggedAndDropped(int slideId, int newSlideIndex)
        {
            if (getCustomSlideById(slideId) != null)
            {
                getCustomSlideById(slideId).updateSlideIndex(newSlideIndex);
            }
        }

        /*
         * Check button clicked.
         * Check if pushed questions will ever be evaluated, if they are pushed/evaluated in the given order.
         */
        private void check_button_Click(object sender, RibbonControlEventArgs e)
        {
            checkQuestionsPushEvaluationOrder(true);
        }

        /// <summary>
        /// Check if pushed questions will ever be evaluated, if they are pushed/evaluated in the given order.
        /// </summary>
        /// <param name="explicitCheck">true if check button is clicked</param>
        /// <returns></returns>
        public bool checkQuestionsPushEvaluationOrder(bool explicitCheck)
        {
            List<String> errorMessages = new List<String>();
            foreach (var customSlide in questionSlides)
            {
                foreach (var question in customSlide.questionList)
                {

                    // question will never be evaluated because no slide is set to evaluate it
                    if (question.EvaluateSlideId == null)
                    {
                        // only add this error message when checking by clicked check button.
                        if (explicitCheck)
                        {
                            errorMessages.Add("Question '" + question.Content + "' pushed on slide number " + customSlide.SlideIndex + " will " +
                                "never be evaluated because you didn't define a slide to evaluate it.");
                        }
                        continue;
                    }

                    // question will never be evaluated because evaluationIndex <= pushIndex
                    if (pptNavigator.getSlideById(question.EvaluateSlideId).SlideIndex <= question.PushSlideIndex)
                    {
                        errorMessages.Add("Question '" + question.Content + "' pushed on slide number " + question.PushSlideIndex + " will " +
                            "never be evaluated because you try to evaluate it on a previous or same slide number (slide number: " +
                            pptNavigator.getSlideById(question.EvaluateSlideId).SlideIndex + ").");
                    }
                }
            }

            // if there is any errorMessage than display popup
            if (errorMessages.Count > 0)
            {
                String errMessage = "";
                foreach (var message in errorMessages)
                {
                    errMessage = errMessage + message + "\n\n";
                }
                MessageBox.Show(errMessage, "You have pushed some questions which will never be evaluated.");
                return false;
            }
            else
            {
                if (explicitCheck)
                {
                    MessageBox.Show("Your push/evaluation order of your question is fine !!!");
                }
                return true;
            }
        }

        /*
         * Start presentation with configured slides.
         */
        private void startSurveyButton_Click(object sender, RibbonControlEventArgs e)
        {
            // check if questions have correct push/evaluation order.
            if (checkQuestionsPushEvaluationOrder(true))
            {
                pptNavigator.startPresentation();
            }
        }

        /*
         * Push questions for given slide id.
         */
        public void pushQuestions(int customSlideId)
        {
            if (getCustomSlideById(customSlideId) != null)
            {
                foreach (var question in getCustomSlideById(customSlideId).questionList)
                {
                    // TODO:
                    // myRestHelper.pushQuestion(question);
                }
            }
        }

        /*
         * This method will open a window during presentation mode to display the evaluation of a 
         * previous pushed question.
         */
        public void evaluateQuestions(int currentSlideId)
        {
            List<String> questions = new List<String>();
            foreach (var customSlide in questionSlides)
            {
                foreach (var question in customSlide.questionList)
                {
                    // TODO: get results from REST
                    questions.Add(question.Content);
                }
            }

            // TODO: open window to display chart
            EvaluationChartForm evaluationForm = new EvaluationChartForm("fubaaaaa");
            evaluationForm.Show();
        }
    }
}
