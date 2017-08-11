using PowerPointAddIn1.utils;
using System;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public partial class StartSessionForm : Form
    {
        // basically to access pptNavigator.
        MyRibbon myRibbon;

        // SessionController to start Session.
        SessionController sessionController;

        // if start session from first slide or not.
        bool fromBeginning;

        /*
         * Constructor.
         */
        public StartSessionForm(MyRibbon myRibbon, bool fromBeginning)
        {
            InitializeComponent();
            this.myRibbon = myRibbon;
            this.fromBeginning = fromBeginning;
            this.sessionController = new SessionController(myRibbon, myRibbon.pptNavigator.pptApplication);
            fillComboboxes();
            start_session_lectures_combo.DisplayMember = "Text";
            start_session_lectures_combo.ValueMember = "Value";
            start_session_chapters_combo.DisplayMember = "Text";
            start_session_chapters_combo.ValueMember = "Value";
        }

        /*
         * Fill lecture and chapter comboboxes with values.
         */
        public void fillComboboxes()
        {
            // if professor has any lectures than fill comboboxes
            if (myRibbon.lectureDropDown.Items.Count > 0)
            {
                // iterate lectureDropDown from myRibbon and add the to this lecture combo
                foreach (var lectureItem in myRibbon.lectureDropDown.Items)
                {
                    start_session_lectures_combo.Items.Add(
                        new { Text = lectureItem.Label, Value = lectureItem.Tag });
                }

                // fill chapter combobox
                String selectedLectureId = (String)(start_session_lectures_combo.Items[0] as dynamic).Value;
                Lecture lecture = myRibbon.getLectureById(selectedLectureId);
                foreach (var chapterItem in lecture.getChapters())
                {
                    start_session_chapters_combo.Items.Add(new { Text = chapterItem.Name, Value = chapterItem.ID });
                }
            }
        }

        /*
         * Is called whenever the selection of lecture combobox has changed.
         */
        public void lectureCombo_SelectionChanged(object sender, EventArgs e)
        {
            start_session_chapters_combo.Items.Clear();

            String selectedLectureId = (String)(start_session_lectures_combo.SelectedItem as dynamic).Value;
            Lecture lecture = myRibbon.getLectureById(selectedLectureId);

            // fill lecture combobox
            foreach (var chapterItem in lecture.getChapters())
            {
                start_session_chapters_combo.Items.Add(new { Text = chapterItem.Name, Value = chapterItem.ID });
            }
        }

        /*
         * Start presentation and record it.
         */
        private void start_session_start_record_button_Click(object sender, EventArgs e)
        {
            if (start_session_lectures_combo.SelectedItem == null ||
                start_session_chapters_combo.SelectedItem == null) {
                start_session_error.Visible = true;
                return;
            }
            String selectedLectureId = (String)(start_session_lectures_combo.SelectedItem as dynamic).Value;
            String selectedChapterId = (String)(start_session_chapters_combo.SelectedItem as dynamic).Value;
            sessionController.startPresentation(fromBeginning, myRibbon.pptNavigator.SlideIndex,
                                                myRibbon.pptNavigator.presentation, myRibbon.pptNavigator.slides,
                                                selectedLectureId, selectedChapterId);
            Close();
        }

        /*
         * Clicked on dont Record button -> Start presentation without recording it.
         */
        private void dontRecordButton_Click(object sender, EventArgs e)
        {
            sessionController.startPresentation(fromBeginning, myRibbon.pptNavigator.SlideIndex,
                                                myRibbon.pptNavigator.presentation, myRibbon.pptNavigator.slides,
                                                null, null);
        }
    }
}
