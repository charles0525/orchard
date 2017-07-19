using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Acai.NewsStatistics.Models
{
    [Serializable]
    public class ListContentsModel
    {
        public int Id { get; set; }
        public string PublishTime { get; set; }
        public string Title { get; set; }
        public string Keywords { get; set; }
        public string LinkUrl { get; set; }
        public string Tags { get; set; }
    }
}