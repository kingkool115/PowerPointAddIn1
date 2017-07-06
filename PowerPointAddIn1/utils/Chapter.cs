using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddIn1
{
    public class Chapter
    {
        private List<Survey> surveyList;

        [JsonProperty(PropertyName = "id")]
        public string ID { get; set; }

        [JsonProperty(PropertyName = "name")]
        public string Name { get; set; }

        public List<Survey> getSurveys()
        {
            return surveyList;
        }

        public void setSurveys(List<Survey> surveyList)
        {
            this.surveyList = surveyList;
        }
    }
}
