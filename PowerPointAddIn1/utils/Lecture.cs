using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddIn1
{
    public class Lecture
    {
        List<Chapter> chapterList;

        [JsonProperty(PropertyName = "id")]
        public string ID { get; set; }

        [JsonProperty(PropertyName = "name")]
        public string Name { get; set; }

        public List<Chapter> getChapters()
        {
            return chapterList;
        }

        public void setChapters(List<Chapter> chapters)
        {
            this.chapterList = chapters;
        }
    }
}
