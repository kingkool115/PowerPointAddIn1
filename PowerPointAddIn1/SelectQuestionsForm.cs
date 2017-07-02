using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
// add PowerPoint namespace
using PPt = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;

namespace PowerPointAddIn1
{
    public partial class SelectQuestionsForm : Form
    {
        MyRibbon myRibbon;
        // Define PowerPoint Application object
        PPt.Application pptApplication;
        // Define Presentation object
        PPt.Presentation presentation;
        // Define Slide collection
        PPt.Slides slides;
        PPt.Slide slide;

        // Slide count
        int slidescount;
        // slide index
        int slideIndex;

        public SelectQuestionsForm()
        {
            InitializeComponent();
            myRibbon = Globals.Ribbons.Ribbon;
            initLecturesCombo();
            // Get Running PowerPoint Application object
            pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;

            if (pptApplication != null)
            {
                // Get Presentation Object
                presentation = pptApplication.ActivePresentation;
                // Get Slide collection object
                slides = presentation.Slides;
                // Get Slide count
                slidescount = slides.Count;
                // Get current selected slide 
                try
                {
                    // Get selected slide object in normal view
                    slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
                }
                catch
                {
                    // set first slide as selected slide
                    if (slides.Count > 0)
                    {
                        slide = slides[1];
                    }
                }
            }
        }

        /*
         * Is called when selectQuestionsForm appears.
         */
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        /*
         * Fill lectures combobox with items.
         */
        public void initLecturesCombo()
        {
            var response = myRibbon.myRestHelper.getAllLectures();

            // fill lectures combobox
            var lectureList = new List<Lecture>();
            lectureList = JsonConvert.DeserializeObject<List<Lecture>>(response.Content);
            lectureComboAddQuestion.DataSource = lectureList;
            lectureComboAddQuestion.DisplayMember = "Name";
            lectureComboAddQuestion.ValueMember = "ID";
        }

        /*
         * Update chapters combobox wheneve lectures combobox is changed.
         */
        public void comboLectures_selectionChanged(object sender, EventArgs e)
        {
            // get chapters of lecture from REST-Service
            Lecture selectedLecture = (Lecture) lectureComboAddQuestion.SelectedItem;
            var response = myRibbon.myRestHelper.getChaptersOfLecture(selectedLecture.ID);

            // fill chapter combobox
            var chapterList = new List<Chapter>();
            chapterList = JsonConvert.DeserializeObject<List<Chapter>>(response.Content);
            chapterComboAddQuestion.DataSource = chapterList;
            chapterComboAddQuestion.DisplayMember = "Name";
            chapterComboAddQuestion.ValueMember = "ID";
        }

        /*
         * Update surveys combobox wheneve chapters combobox is changed.
         */
        public void comboChapters_selectionChanged(object sender, EventArgs e)
        {
            // execute the request
            Lecture selectedLecture = (Lecture) lectureComboAddQuestion.SelectedItem;
            Chapter selectedChapter = (Chapter) chapterComboAddQuestion.SelectedItem;
            var response = myRibbon.myRestHelper.getSurveysOfChapter(selectedLecture.ID, selectedChapter.ID);

            // fill surveys combobox
            var surveysList = new List<Survey>();
            surveysList = JsonConvert.DeserializeObject<List<Survey>>(response.Content);
            surveyComboAddQuestion.DataSource = surveysList;
            surveyComboAddQuestion.DisplayMember = "Name";
            surveyComboAddQuestion.ValueMember = "ID";
        }

        public void comboSurveys_selectionChanged(object sender, EventArgs e)
        {
            // clear listview
            questionsListView.Items.Clear();

            // execute the request
            Lecture selectedLecture = (Lecture) lectureComboAddQuestion.SelectedItem;
            Chapter selectedChapter = (Chapter) chapterComboAddQuestion.SelectedItem;
            Survey selectedSurvey = (Survey) surveyComboAddQuestion.SelectedItem;
            var response = myRibbon.myRestHelper.getQuestionsOfSurvey(selectedLecture.ID, selectedChapter.ID, selectedSurvey.ID);

            // fill surveys combobox
            var questionList = new List<Question>();
            questionList = JsonConvert.DeserializeObject<List<Question>>(response.Content);

            // TODO: fill listview with questions
            foreach (var question in questionList)
            {
                // item is row
                ListViewItem row = new ListViewItem(question.Content);
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
                questionsListView.Items.Add(row);
            }
        }

        /*
         * Handles click on Add-Button.
         */
        private void addQuestion_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in questionsListView.Items)
            {
                if (item.Checked)
                {
                    // add to other list.
                    Lecture lecture =  (Lecture) lectureComboAddQuestion.SelectedItem;
                    Chapter chapter = (Chapter) chapterComboAddQuestion.SelectedItem;
                    Survey survey = (Survey) surveyComboAddQuestion.SelectedItem;
                    String question = item.Text;

                    ListViewItem row = new ListViewItem(question);
                    row.SubItems.Add(lecture.Name);
                    row.SubItems.Add(chapter.Name);
                    row.SubItems.Add(survey.Name);
                    listView1.Items.Add(row);
                }
            }
        }

        /*
         * Handles remove Button.
         */
        private void removeQuestionsButton_Click(object sender, EventArgs e)
        {

        }

        /*
         * Handles save questions button.
         */
        private void saveQuestionsButton_click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void nextSlideButton_Click(object sender, EventArgs e)
        {
            slideIndex = slide.SlideIndex + 1;
            if (slideIndex > slidescount)
            {
                MessageBox.Show("It is already last page");
            }
            else
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    pptApplication.SlideShowWindows[1].View.Next();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
        }

        private void previousSlideButton_Click(object sender, EventArgs e)
        {
            slideIndex = slide.SlideIndex - 1;
            if (slideIndex >= 1)
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    pptApplication.SlideShowWindows[1].View.Previous();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
            else
            {
                MessageBox.Show("It is already Fist Page");
            }
        }
    }
}
