using Orchard.ContentManagement;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Acai.NewsStatistics.Models
{
    public class ListContentsViewModel
    {
        public ListContentsViewModel()
        {
            Options = new ContentOptions();
        }

        public string Id { get; set; }

        public string TypeName
        {
            get { return Id; }
        }

        public string TypeDisplayName { get; set; }
        public int? Page { get; set; }
        public IList<Entry> Entries { get; set; }
        public ContentOptions Options { get; set; }

        #region Nested type: Entry

        public class Entry
        {
            public ContentItem ContentItem { get; set; }
            public ContentItemMetadata ContentItemMetadata { get; set; }
        }

        #endregion

        public string Code { get; set; }
        public string PublishTime { get; set; } = DateTime.Now.AddDays(-7).ToString("yyyy-MM-dd");
        public string Keywords { get; set; }
        public string Tags { get; set; }
    }

    public class ContentOptions
    {
        public ContentOptions()
        {
            OrderBy = ContentsOrder.Modified;
            BulkAction = ContentsBulkAction.None;
            ContentsStatus = ContentsStatus.Latest;
        }
        public string SelectedFilter { get; set; }
        public string SelectedCulture { get; set; }
        public IEnumerable<KeyValuePair<string, string>> FilterOptions { get; set; }
        public ContentsOrder OrderBy { get; set; }
        public ContentsStatus ContentsStatus { get; set; }
        public ContentsBulkAction BulkAction { get; set; }
        public IEnumerable<string> Cultures { get; set; }
    }

    public enum ContentsOrder
    {
        Modified,
        Published,
        Created
    }

    public enum ContentsStatus
    {
        Draft,
        Published,
        AllVersions,
        Latest,
        Owner
    }

    public enum ContentsBulkAction
    {
        None,
        PublishNow,
        Unpublish,
        Remove
    }
}