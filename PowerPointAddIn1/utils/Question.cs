using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace PowerPointAddIn1
{
    public class Question
    {
        private Lecture lecture;
        private Chapter chapter;
        private Survey survey;
        public int? PushSlideId {get; set;}
        public int? EvaluateSlideId { get; set; }
        public int? PushSlideIndex {get; set;}

        [JsonProperty(PropertyName = "id")]
        public string ID { get; set; }

        [JsonProperty(PropertyName = "question")]
        public string Content { get; set; }

        [JsonProperty(PropertyName = "is_text_response")]
        public int isTextResponse { get; set; }

        public void setLectureChapterSurvey(Lecture lecture, Chapter chapter, Survey survey)
        {
            this.lecture = lecture;
            this.chapter = chapter;
            this.survey = survey;
        }

        public Lecture getLecture()
        {
            return lecture;
        }

        public Chapter getChapter()
        {
            return chapter;
        }

        public Survey getSurvey()
        {
            return survey;
        }
    }
}
