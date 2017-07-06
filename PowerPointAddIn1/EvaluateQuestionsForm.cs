using PowerPointAddIn1.utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public partial class EvaluateQuestionsForm : Form
    {
        MyRibbon myRibbon;
        List<Question> notEvaluatedQuestionsList;
        List<Question> evaluatedQuestionsList;

        public EvaluateQuestionsForm()
        {
            InitializeComponent();
            myRibbon = Globals.Ribbons.Ribbon;
            notEvaluatedQuestionsList = new List<Question>();
            evaluatedQuestionsList = new List<Question>();
        }

        private void SelectAnswersForm_Load(object sender, EventArgs e)
        {
            updateListViews();
        }

        /*
         * Update the content of both list views.
         */
        private void updateListViews()
        {
            // clear list views
            this.notEvaluatedQuestionsListView.Items.Clear();
            this.evaluateQuestionsListView.Items.Clear();

            // clear lists
            evaluatedQuestionsList.Clear();
            notEvaluatedQuestionsList.Clear();

            int currentSlideIndex = myRibbon.pptNavigator.getCurrentSlide();
            
            List<CustomSlide> customSlides = myRibbon.questionSlides;

            // iterate though all custom slides and their questions to decide
            // in which listview a question should be added
            foreach (var cs in customSlides)
            {
                foreach (var question in cs.getQuestions())
                {
                    // item is row
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

                    row.SubItems.Add(cs.getSlideIndex().ToString());
                    row.SubItems.Add(isMultipleChoice);
                    row.SubItems.Add(question.getLecture().Name);
                    row.SubItems.Add(question.getChapter().Name);
                    row.SubItems.Add(question.getSurvey().Name);

                    // if question slide index (when pushed) < current slide index and if question was not evaluated yet
                    // fill not evaluatedQuestionListView
                    if (cs.getSlideIndex() < currentSlideIndex && question.EvaluateSlideIndex == null)
                    {
                        // add to listview
                        notEvaluatedQuestionsListView.Items.Add(row);
                        // add to list
                        notEvaluatedQuestionsList.Add(question);
                    }
                    // if question has a evaluatedSlideIndex -> add to evaluateQuestionsListView
                    if (question.EvaluateSlideIndex == currentSlideIndex)
                    {
                        // add to listview
                        evaluateQuestionsListView.Items.Add(row);
                        // add to list
                        evaluatedQuestionsList.Add(question);
                    }
                }
            } 
        }

        /*
         * Next Button is clicked to move to next slide.
         */
        private void nextSlideButton_Click(object sender, EventArgs e)
        {
            myRibbon.pptNavigator.nextSlide();
            int slideIndex = myRibbon.pptNavigator.getCurrentSlide();
            this.Text = "Select a question to evaluate on slide number " + slideIndex;
            updateListViews();
        }

        /*
         * Previous Button is clicked to move to previoues slide.
         */
        private void previousSlideButton_Click(object sender, EventArgs e)
        {
            myRibbon.pptNavigator.previousSlide();
            int slideIndex = myRibbon.pptNavigator.getCurrentSlide();
            this.Text = "Select a question to evaluate on slide number " + slideIndex;
            updateListViews();
        }

        /*
         * Get a certain question from notEvaluatedQuestionsListView.
         */
        private Question getQuestionFromNotEvaluatedQuestionsListView(String questionId)
        {
            foreach (var question in notEvaluatedQuestionsList)
            {
                if (question.ID.Equals(questionId))
                {
                    return question;
                }
            }
            return null;
        }

        /*
         * Get a certain question from evaluatedQuestionsListView.
         */
        private Question getQuestionFromEvaluatedQuestionsListView(String questionId)
        {
            foreach (var question in evaluatedQuestionsList)
            {
                if (question.ID.Equals(questionId))
                {
                    return question;
                }
            }
            return null;
        }

        /*
         * /*
         * Handles click of Add-Evaluation-Button.
         */
        private void evaluateQuestionsButton_Click(object sender, EventArgs e)
        {
            int slideIndex = myRibbon.pptNavigator.getCurrentSlide();
            foreach (ListViewItem item in notEvaluatedQuestionsListView.Items)
            {
                if (item.Checked)
                {
                    // get question instance
                    Question question = getQuestionFromNotEvaluatedQuestionsListView((String)item.Tag);
                    
                    // add to myRibbon.questionSlides
                    myRibbon.addEvaluationToSlide(slideIndex, question);
                }
            }
            updateListViews();
        }

        /*
         * Handles click of Remove-Evaluation-Button.
         */
        private void removeQuestionEvaluationButton_Click(object sender, EventArgs e)
        {
            int slideIndex = myRibbon.pptNavigator.getCurrentSlide();
            foreach (ListViewItem item in evaluateQuestionsListView.Items)
            {
                if (item.Checked)
                {
                    // get question instance
                    Question question = getQuestionFromEvaluatedQuestionsListView((String)item.Tag);

                    // add to myRibbon.questionSlides
                    myRibbon.removeEvaluationFromSlide(slideIndex, question);
                }
            }
            updateListViews();
        }
    }
}
