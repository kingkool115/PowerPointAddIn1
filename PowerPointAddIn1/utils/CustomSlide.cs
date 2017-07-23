using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Newtonsoft.Json;

namespace PowerPointAddIn1.utils
{
    [JsonObject]
    public class CustomSlide
    {
        [JsonProperty(PropertyName = "slide_index")]
        public int SlideIndex { get; set; }
        [JsonProperty(PropertyName = "slide_id")]
        public int SlideId { get; }
        [JsonProperty(PropertyName = "question_list")]
        public List<QuestionObj> questionList { get; set; }

        public CustomSlide(int slideId, int slideIndex,  QuestionObj question)
        {
            this.SlideId = slideId;
            this.SlideIndex = slideIndex;
            questionList = new List<QuestionObj>();
            question.PushSlideId = slideId;
            questionList.Add(question);
        }

        [JsonConstructor]
        public CustomSlide(int slideId, int slideIndex, List<QuestionObj> questionList)
        {
            this.SlideId = slideId;
            this.SlideIndex = slideIndex;
            this.questionList = questionList;
        }

        public QuestionObj getQuestion(QuestionObj question)
        {
            foreach (var qu in questionList)
            {
                if (question == qu)
                {
                    return qu;
                }
            }
            return null;
        }

        /*
         * This method is called whenever slides are added/removed to current presentation
         */
        public void updateSlideIndex(int newSlideIndex)
        {
            // update SlideIndex of CustomSlide
            SlideIndex = newSlideIndex;

            // update PushSlideIndex of all its questions
            foreach (var question in questionList)
            {
                question.PushSlideIndex = SlideIndex;
            }
        }

        /*
         * Add a question to this slide.
         */
        public void addQuestion(QuestionObj question)
        {
            if (!questionExists(question))
            {
                question.PushSlideIndex = SlideIndex;
                question.PushSlideId = SlideId;
                questionList.Add(question);
            }
        }

        /*
         * Remove a question from this slide. 
         */
        public void removeQuestion(QuestionObj question)
        {
            foreach (var qu in questionList)
            {
                if (qu.ID.Equals(question.ID))
                {
                    qu.PushSlideIndex = null;
                    qu.PushSlideId = null;
                    questionList.Remove(qu);
                    break;
                }
            }
        }

        /*
         * Checks if the question does already exists for this slide.
         */
        private Boolean questionExists(QuestionObj question)
        {
            foreach (var qu in questionList)
            {
                if (qu.ID.Equals(question.ID))
                {
                    return true;
                }
            }
            return false;
        }
        
    }
}
