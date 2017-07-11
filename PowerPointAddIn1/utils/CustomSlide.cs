using System;
using System.Collections.Generic;
using Microsoft.Office.Core;

namespace PowerPointAddIn1.utils
{
    public class CustomSlide
    {
        public int SlideIndex { get; set; }
        public int SlideId { get; }
        List<Question> questionList;

        public CustomSlide(int slideId, int slideIndex,  Question question)
        {
            this.SlideId = slideId;
            this.SlideIndex = slideIndex;
            questionList = new List<Question>();
            question.PushSlideId = slideId;
            questionList.Add(question);
        }

        public Question getQuestion(Question question)
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
         * Returns all questions of this slide.
         */
        public List<Question> getQuestions()
        {
            return questionList;
        }

        /*
         * Add a question to this slide.
         */
        public void addQuestion(Question question)
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
        public void removeQuestion(Question question)
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
        private Boolean questionExists(Question question)
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
