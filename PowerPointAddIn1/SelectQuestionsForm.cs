using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public partial class SelectQuestionsForm : Form
    {
        MyRibbon myRibbon;
        List<QuestionObj> possibleQuestionsList;
        List<QuestionObj> questionsForCurrentSlide;

        public SelectQuestionsForm()
        {
            InitializeComponent();
            myRibbon = Globals.Ribbons.Ribbon;
            initLecturesCombo();
        }

        private void Form1_Load(object sender, EventArgs ea)
        {
            this.Text = "Select a question to add to slide number " + myRibbon.pptNavigator.SlideIndex;
        }

        /*
         * Fill lectures combobox with items.
         */
        public void initLecturesCombo()
        {
            // fill lectures combobox
            var lectureList = myRibbon.myLectures;
            lectureComboAddQuestion.DataSource = lectureList;
            lectureComboAddQuestion.DisplayMember = "Name";
            lectureComboAddQuestion.ValueMember = "ID";
        }

        /*
         * Update chapters combobox wheneve lectures combobox is changed.
         */
        public void comboLectures_selectionChanged(object sender, EventArgs e)
        {
            Lecture selectedLecture = (Lecture) lectureComboAddQuestion.SelectedItem;

            // fill chapter combobox
            var chapterList = selectedLecture.getChapters();
            chapterComboAddQuestion.DataSource = chapterList;
            chapterComboAddQuestion.DisplayMember = "Name";
            chapterComboAddQuestion.ValueMember = "ID";
        }

        /*
         * Update surveys combobox whenever chapters combobox is changed.
         */
        public void comboChapters_selectionChanged(object sender, EventArgs e)
        {
            Chapter selectedChapter = (Chapter) chapterComboAddQuestion.SelectedItem;

            // fill surveys combobox
            var surveysList = selectedChapter.getSurveys();
            surveyComboAddQuestion.DataSource = surveysList;
            surveyComboAddQuestion.DisplayMember = "Name";
            surveyComboAddQuestion.ValueMember = "ID";
        }

        /*
         * Update possibleQuestionsListView whenever surveys combobox is changed.
         */
        public void comboSurveys_selectionChanged(object sender, EventArgs e)
        {
            // clear listview
            possibleQuestionsListView.Items.Clear();
            
            Lecture selectedLecture = (Lecture) lectureComboAddQuestion.SelectedItem;
            Chapter selectedChapter = (Chapter) chapterComboAddQuestion.SelectedItem;
            Survey selectedSurvey = (Survey) surveyComboAddQuestion.SelectedItem;

            // fill surveys combobox
            possibleQuestionsList = selectedSurvey.getQuestions();

            // fill listview with questions
            foreach (var question in possibleQuestionsList)
            {
                // item is row
                ListViewItem row = new ListViewItem(question.Content);
                row.Tag = question.ID;

                String isMultipleChoice;
                if (question.isTextResponse == 1)
                {
                    isMultipleChoice = "no";
                } else
                {
                    isMultipleChoice = "yes";
                }

                row.SubItems.Add(isMultipleChoice);
                // subitem represents column
                possibleQuestionsListView.Items.Add(row);
            }
        }

        /*
         * Get a certain question from possibleQuestionsListView.
         */
        private QuestionObj getQuestionFromPossibleQuestionsListView(String questionId)
        {
            foreach (var question in possibleQuestionsList)
            {
                if (question.ID.Equals(questionId))
                {
                    return question;
                }
            }
            return null;
        }

        /*
         * Get a certain question from questionsPerSlideListView.
         */
        private QuestionObj getQuestionFromQuestionsPerSlideListView(String questionId)
        {
            foreach (var question in questionsForCurrentSlide)
            {
                if (question.ID.Equals(questionId))
                {
                    return question;
                }
            }
            return null;
        }

        /*
         * Handles click on Add-Button.
         */
        private void addQuestion_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in possibleQuestionsListView.Items)
            {
                if (item.Checked)
                {
                    // get question instance
                    QuestionObj question = getQuestionFromPossibleQuestionsListView((String) item.Tag);

                    // add to myRibbon.questionSlides
                    int slideIndex = myRibbon.pptNavigator.SlideIndex;
                    int slideId = myRibbon.pptNavigator.SlideId;
                    myRibbon.addQuestionToSlide(slideId, slideIndex, question);

                }
            }
            updateQuestionsPerSlideListView();
        }

        /*
         * Handles remove Button.
         */
        private void removeQuestionsButton_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in questionsPerSlideListView.Items)
            {
                if (item.Checked)
                {
                    // get question instance
                    QuestionObj question = getQuestionFromQuestionsPerSlideListView((String)item.Tag);

                    // add to myRibbon.questionSlides
                    myRibbon.removeQuestionFromSlide(myRibbon.pptNavigator.SlideId, question);

                }
            }
            updateQuestionsPerSlideListView();
        }

        /*
         * Is called when Add-button, Remove-button is clicked or whenever current slide has changed.
         */
        public void updateQuestionsPerSlideListView()
        {
            // display all questions of that slide in questionsPerSlideListView
            questionsPerSlideListView.Items.Clear();
            if (myRibbon.getCustomSlideById(myRibbon.pptNavigator.SlideId) != null)
            {
                questionsForCurrentSlide = myRibbon.getCustomSlideById(myRibbon.pptNavigator.SlideId).questionList;
                foreach (var question in questionsForCurrentSlide)
                {
                    ListViewItem row = new ListViewItem(question.Content);
                    row.Tag = question.ID;
                    String isMultipleChoice;
                    if (question.isTextResponse == 1)
                    {
                        isMultipleChoice = "no";
                    }
                    else
                    {
                        isMultipleChoice = "yes";
                    }
                    row.SubItems.Add(isMultipleChoice);
                    row.SubItems.Add(question.Lecture.Name);
                    row.SubItems.Add(question.Chapter.Name);
                    row.SubItems.Add(question.Survey.Name);
                    questionsPerSlideListView.Items.Add(row);
                }
            }
        }

        /*
         * Handles save questions button.
         */
        private void saveQuestionsButton_click(object sender, EventArgs e)
        {

        }

        /*
         * Next Button is clicked to move to next slide.
         */
        private void nextSlideButton_Click(object sender, EventArgs e)
        {
            myRibbon.pptNavigator.nextSlide();
            this.Text = "Select a question to add to slide number " + myRibbon.pptNavigator.SlideIndex;
        }

        /*
         * Previous Button is clicked to move to previoues slide.
         */
        private void previousSlideButton_Click(object sender, EventArgs e)
        {
            myRibbon.pptNavigator.previousSlide();
            this.Text = "Select a question to add to slide number " + myRibbon.pptNavigator.SlideIndex;
        }
    }
}
