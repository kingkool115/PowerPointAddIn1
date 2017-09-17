using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public partial class SelectLectureForm : Form
    {

        // basically to access pptNavigator.
        MyRibbon myRibbon;
        
        // list of lectures
        List<Lecture> lectureList;

        public SelectLectureForm(MyRibbon myRibbon)
        {

            InitializeComponent();
            this.myRibbon = myRibbon;
            fillCombobox();
            select_lectures_combo.DisplayMember = "Text";
            select_lectures_combo.ValueMember = "Value";
        }

        /*
         * Fill lecture combobox with values.
         */
        public void fillCombobox()
        {
            // fill lecture combo with all available lectures
            var response = myRibbon.myRestHelper.getAllAvailableLectures();
            var content = response.Content;
            lectureList = JsonConvert.DeserializeObject<List<Lecture>>(content);
            foreach (var lectureItem in lectureList)
            {
                select_lectures_combo.Items.Add(
                    new { Text = lectureItem.Name, Value = lectureItem.ID });
            }

            // fill lectureList
            foreach (var lect in lectureList)
            {
                var chapterList = myRibbon.myRestHelper.GetChaptersOfLectureAsGuest(lect.ID);
                lect.setChapters(chapterList);
            }

            // set selected lecture for this presentation
            if (myRibbon.LectureForThisPresentation != null)
            {
                select_lectures_combo.SelectedItem =
                    new { Text = myRibbon.LectureForThisPresentation.Name, Value = myRibbon.LectureForThisPresentation.ID };
            }
        }

        /*
         * Select Lecture.
         */
        private void select_lecture_button_Click(object sender, EventArgs e)
        {
            String selectedLectureId = (String)(select_lectures_combo.SelectedItem as dynamic).Value;
            Lecture lecture = getLectureById(selectedLectureId);
            myRibbon.LectureForThisPresentation = lecture;
            myRibbon.select_lecture_group.Label = "     Lecture: " + lecture.Name + "     ";
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
