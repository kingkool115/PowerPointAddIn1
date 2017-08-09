using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddIn1.utils
{
    public class Evaluation
    {
        [JsonProperty(PropertyName = "question_id")]
        public String QuestionId { get; set; }
        [JsonProperty(PropertyName = "question")]
        public String Question { get; set; }
        [JsonProperty(PropertyName = "is_text_response")]
        public Boolean IsTextResponse { get; set; }
        [JsonProperty(PropertyName = "answers")]
        public Dictionary<String, int> Answers { get; set; }
    }
}
