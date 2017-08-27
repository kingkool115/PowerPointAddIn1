using Newtonsoft.Json;
using PowerPointAddIn1.utils;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public partial class StartSessionForm : Form
    {
        // basically to access pptNavigator.
        MyRibbon myRibbon;
        
        // if start session from first slide or not.
        bool fromBeginning;

        // if someone starts the presentation who didn't log in before, than false
        bool isUserLoggedIn;

        // list of lectures
        List<Lecture> lectureList;

        /*
         * Constructor.
         */
        public StartSessionForm(MyRibbon myRibbon, bool fromBeginning, bool isUserLoggedIn)
        {
            InitializeComponent();
            this.myRibbon = myRibbon;
            this.fromBeginning = fromBeginning;
            this.isUserLoggedIn = isUserLoggedIn;
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
            // fill lecture combo with all available lectures
            var response = myRibbon.myRestHelper.getAllAvailableLectures();
            var content = response.Content;
            lectureList = JsonConvert.DeserializeObject<List<Lecture>>(content);
            foreach (var lectureItem in lectureList)
            {
                start_session_lectures_combo.Items.Add(
                    new { Text = lectureItem.Name, Value = lectureItem.ID });
            }
            // fill lectureList
            foreach (var lect in lectureList)
            {
                var chapterList = myRibbon.myRestHelper.GetChaptersOfLectureAsGuest(lect.ID);
                lect.setChapters(chapterList);
            }

            // fill chapter combobox
            String selectedLectureId = (String)(start_session_lectures_combo.Items[0] as dynamic).Value;
            Lecture lecture = getLectureById(selectedLectureId);
            foreach (var chapterItem in lecture.getChapters())
            {
                start_session_chapters_combo.Items.Add(new { Text = chapterItem.Name, Value = chapterItem.ID });
            }
        }

        /*
         * Is called whenever the selection of lecture combobox has changed -> update chapter combobox.
         */
        public void lectureCombo_SelectionChanged(object sender, EventArgs e)
        {
            start_session_chapters_combo.Items.Clear();

            String selectedLectureId = (String)(start_session_lectures_combo.SelectedItem as dynamic).Value;
            Lecture lecture = getLectureById(selectedLectureId);

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
            myRibbon.startNewSession(selectedLectureId, selectedChapterId, fromBeginning);
            Close();
        }

        /*
         * Get a lecture from myLectures by id.
         */
        public Lecture getLectureById(String lectureId)
        {
            foreach (var lecture in lectureList)
            {
                if (lecture.ID.Equals(lectureId))
                {
                    return lecture;
                }
            }
            return null;
        }

    }
}
