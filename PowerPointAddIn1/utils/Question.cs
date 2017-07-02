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
        [JsonProperty(PropertyName = "id")]
        public string ID { get; set; }

        [JsonProperty(PropertyName = "question")]
        public string Content { get; set; }

        [JsonProperty(PropertyName = "is_text_response")]
        public int isTextResponse { get; set; }
    }
}
