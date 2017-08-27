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
        public List<Question> PushQuestionList { get; set; }
        [JsonProperty(PropertyName = "eval_question_list")]
        public List<Question> EvaluationList { get; set; }

        /*
         * Constructor. 
         */
        public CustomSlide(int slideId, int slideIndex,  Question question, bool evaluation)
        {
            SlideId = slideId;
            SlideIndex = slideIndex;
            PushQuestionList = new List<Question>();
            EvaluationList = new List<Question>();
            question.PushSlideId = slideId;
            if (evaluation)
            {
                question.EvaluateSlideId = slideId;
                question.EvaluateSlideIndex = slideIndex;
                EvaluationList.Add(question);
            }
            else {
                question.PushSlideIndex = slideIndex;
                question.PushSlideId = slideId;
                PushQuestionList.Add(question);
            }
        }

        [JsonConstructor]
        public CustomSlide(int slideId, int slideIndex, List<Question> PushQuestionList)
        {
            this.SlideId = slideId;
            this.SlideIndex = slideIndex;
            this.PushQuestionList = PushQuestionList;
        }

        public CustomSlide(int slideId, int slideIndex, List<Question> PushQuestionList, List<Question> EvaluationList)
        {
            this.SlideId = slideId;
            this.SlideIndex = slideIndex;
            this.PushQuestionList = PushQuestionList;
            this.EvaluationList = EvaluationList;
        }

        /*
         * Get Push Question from PushQuestionList.
         */
        public Question getPushQuestion(Question question)
        {
            foreach (var qu in PushQuestionList)
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
        public void updatePushSlideIndex(int newSlideIndex)
        {
            // update SlideIndex of CustomSlide
            SlideIndex = newSlideIndex;

            // update PushSlideIndex of all its questions
            foreach (var question in PushQuestionList)
            {
                question.PushSlideIndex = SlideIndex;
            }
        }

        /*
         * Add a push question to this slide.
         */
        public void addPushQuestion(Question question)
        {
            if (!pushQuestionExists(question))
            {
                question.PushSlideIndex = SlideIndex;
                question.PushSlideId = SlideId;
                PushQuestionList.Add(question);
            }
        }

        /*
         * Remove a push question from this slide. 
         */
        public void removePushQuestion(Question question)
        {
            foreach (var qu in PushQuestionList)
            {
                if (qu.ID.Equals(question.ID))
                {
                    qu.PushSlideIndex = null;
                    qu.PushSlideId = null;
                    PushQuestionList.Remove(qu);
                    break;
                }
            }
        }

        /*
         * Checks if the push question does already exists for this slide.
         */
        private Boolean pushQuestionExists(Question question)
        {
            foreach (var qu in PushQuestionList)
            {
                if (qu.ID.Equals(question.ID))
                {
                    return true;
                }
            }
            return false;
        }


        /*
         * Get Evaluation Question from PushQuestionList.
         */
        public Question getEvaluation(Question question)
        {
            foreach (var eval in EvaluationList)
            {
                if (question == eval)
                {
                    return eval;
                }
            }
            return null;
        }

        /*
         * Add a push question to this slide.
         */
        public void addEvaluation(Question question)
        {
            if (!evaluationExists(question))
            {
                question.EvaluateSlideId = SlideId;
                question.EvaluateSlideIndex = SlideIndex;
                EvaluationList.Add(question);
                // update slideId/slideIndex in pushQuestionsList
                foreach (var pushQuestion in PushQuestionList)
                {
                    if (pushQuestion == question)
                    {
                        pushQuestion.EvaluateSlideId = SlideId;
                        pushQuestion.EvaluateSlideIndex = SlideIndex;
                        break;
                    }
                }
            }
        }

        /*
         * Remove a push question from this slide. 
         */
        public void removeEvaluation(Question question)
        {
            foreach (var qu in EvaluationList)
            {
                if (qu.ID.Equals(question.ID))
                {
                    question.EvaluateSlideId = null;
                    question.EvaluateSlideIndex = null;
                    EvaluationList.Remove(qu);
                    // update slideId/slideIndex in pushQuestionsList
                    foreach (var pushQuestion in PushQuestionList)
                    {
                        if (pushQuestion == question)
                        {
                            pushQuestion.EvaluateSlideId = null;
                            pushQuestion.EvaluateSlideIndex = null;
                            break;
                        }
                    }
                    break;
                }
            }
        }

        /*
         * Checks if the push question does already exists for this slide.
         */
        private Boolean evaluationExists(Question question)
        {
            foreach (var qu in EvaluationList)
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
