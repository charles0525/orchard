using Orchard.UI.Resources;

namespace Acai.NewsStatistics
{
    public class ResourceManifest : IResourceManifestProvider {
        public void BuildManifests(ResourceManifestBuilder builder) {
            var manifest = builder.Add();
            //manifest.DefineStyle("Admin").SetUrl("orchard-comments-admin.css");

            //manifest.DefineStyle("Calendars")
            //     .SetUrl("/Modules/Orchard.Resources/styles/Calendars/jquery.calendars.picker.full.min.css");

            //manifest.DefineScript("jquery.plugin")
            //   .SetUrl("/Modules/Orchard.Resources/scripts/jquery.plugin.js").SetDependencies("jQuery");

            //manifest.DefineScript("calendars.all")
            //   .SetUrl("/Modules/Orchard.Resources/scripts/Calendars/jquery.calendars.all.js").SetDependencies("jQuery");

            //manifest.DefineScript("calendars.picker")
            //    .SetUrl("/Modules/Orchard.Resources/scripts/Calendars/jquery.calendars.picker.full.min.js").SetDependencies("jQuery");

            manifest.DefineScript("calendars")
              .SetUrl("Plugins/My97DatePicker/WdatePicker.js", "Plugins/My97DatePicker/WdatePicker.js").SetDependencies("jQuery");

            manifest.DefineScript("Acai.NewsStatistics.List")
               .SetUrl("Acai.NewsStatistics.List.js", "Acai.NewsStatistics.List.js").SetDependencies("jQuery");
        }
    }
}