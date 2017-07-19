using Orchard.Localization;
using Orchard.UI.Navigation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Acai.NewsStatistics
{
    public class AdminMenu : INavigationProvider
    {
        public string MenuName => "admin";

        public Localizer T { get; set; }

        public void GetNavigation(NavigationBuilder builder)
        {

            builder.AddImageSet("Acai.NewsStatistics").Add(T("资讯统计"), "1.4", item =>
            {
                item.Action("List", "Admin", new { area = "Acai.NewsStatistics" });

            });
        }
    }
}