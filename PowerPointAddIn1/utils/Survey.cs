using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddIn1
{
    public class Survey
    {
        private List<Question> questionList;

        [JsonProperty(PropertyName = "id")]
        public string ID { get; set; }

        [JsonProperty(PropertyName = "name")]
        public string Name { get; set; }

        public void setQuestions(List<Question> questionList)
        {
            this.questionList = questionList;
        }

        public List<Question> getQuestions()
        {
            return questionList;
        }
    }
}
