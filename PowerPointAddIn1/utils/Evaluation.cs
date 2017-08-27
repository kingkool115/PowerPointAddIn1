﻿using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace PowerPointAddIn1.utils
{
    public class Evaluation
    {
        [JsonProperty(PropertyName = "question_id")]
        public String QuestionId { get; set; }
        [JsonProperty(PropertyName = "image_path")]
        public String ImageURL { get; set; }
        [JsonProperty(PropertyName = "question")]
        public String Question { get; set; }
        [JsonProperty(PropertyName = "is_text_response")]
        public Boolean IsTextResponse { get; set; }

        // needed if deserializing answers but there are no answers available yet
        public Dictionary<String, int> Answers
        {
            get
            {
                var json = this.AnswersJson.ToString();
                if (json == "[]")
                {
                    return new Dictionary<String, int>();
                }
                else
                {
                    return JsonConvert.DeserializeObject<Dictionary<String, int>>(json);
                }
            }
        }

        [JsonProperty(PropertyName = "answers")]
        public object AnswersJson { get; set; }
    }
}
