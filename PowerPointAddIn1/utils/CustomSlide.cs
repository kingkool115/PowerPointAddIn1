using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddIn1.utils
{
    public class CustomSlide
    {
        int slideIndex;
        List<Question> questionList;

        public CustomSlide(int slideIndex, Question question)
        {
            this.slideIndex = slideIndex;
            questionList = new List<Question>();
            question.PushSlideIndex = slideIndex;
            questionList.Add(question);
        }

        /*
         * Returns slide index of this slide.
         */
        public int getSlideIndex()
        {
            return slideIndex;
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
                question.PushSlideIndex = slideIndex;
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
