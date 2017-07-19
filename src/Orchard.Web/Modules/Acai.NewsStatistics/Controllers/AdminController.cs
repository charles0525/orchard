using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Acai.NewsStatistics.Models;
using Orchard.UI.Navigation;
using Orchard.Settings;
using Orchard;
using Orchard.Localization;
using Orchard.Logging;
using Orchard.ContentManagement;
using Orchard.DisplayManagement;
using Orchard.ContentManagement.MetaData.Models;
using Orchard.ContentManagement.MetaData;
using Orchard.Core.Common.Models;
using Orchard.Localization.Services;
using Acai.NewsStatistics.Common;
using Acai.NewsStatistics.Models.Settings;
using Orchard.ContentManagement.Records;
using System.Xml;

namespace Acai.NewsStatistics.Controllers
{
    public class AdminController : Controller
    {
        private readonly ISiteService _siteService;
        private readonly IContentManager _contentManager;
        private readonly IContentDefinitionManager _contentDefinitionManager;
        private readonly ICultureFilter _cultureFilter;
        private readonly ICultureManager _cultureManager;

        private static List<ExcelColumns> listCols = null;
        private static IEnumerable<ListContentsModel> listExportDatas = null;

        public AdminController(
            IOrchardServices orchardServices,
            IContentManager contentManager,
            IContentDefinitionManager contentDefinitionManager,
            ISiteService siteService, IShapeFactory shapeFactory,
            ICultureFilter cultureFilter,
            ICultureManager cultureManager
            )
        {
            _siteService = siteService;
            _contentManager = contentManager;
            _contentDefinitionManager = contentDefinitionManager;
            _cultureFilter = cultureFilter;
            _cultureManager = cultureManager;
            Services = orchardServices;

            T = NullLocalizer.Instance;
            Logger = NullLogger.Instance;
            Shape = shapeFactory;
        }

        dynamic Shape { get; set; }
        public IOrchardServices Services { get; private set; }
        public Localizer T { get; set; }
        public ILogger Logger { get; set; }

        public ActionResult List(ListContentsViewModel model, PagerParameters pagerParameters)
        {
            if (model == null)
            {
                model = new ListContentsViewModel();
            }

            Pager pager = new Pager(_siteService.GetSiteSettings(), pagerParameters);

            var versionOptions = VersionOptions.Latest;
            switch (model.Options.ContentsStatus)
            {
                case ContentsStatus.Published:
                    versionOptions = VersionOptions.Published;
                    break;
                case ContentsStatus.Draft:
                    versionOptions = VersionOptions.Draft;
                    break;
                case ContentsStatus.AllVersions:
                    versionOptions = VersionOptions.AllVersions;
                    break;
                default:
                    versionOptions = VersionOptions.Latest;
                    break;
            }

            var query = _contentManager.Query(versionOptions, GetListableTypes(false).Select(ctd => ctd.Name).ToArray());

            if (!string.IsNullOrEmpty(model.TypeName))
            {
                var contentTypeDefinition = _contentDefinitionManager.GetTypeDefinition(model.TypeName);
                if (contentTypeDefinition == null)
                    return HttpNotFound();

                model.TypeDisplayName = !string.IsNullOrWhiteSpace(contentTypeDefinition.DisplayName)
                                            ? contentTypeDefinition.DisplayName
                                            : contentTypeDefinition.Name;

                // We display a specific type even if it's not listable so that admin pages
                // can reuse the Content list page for specific types.
                query = _contentManager.Query(versionOptions, model.TypeName);
            }

            switch (model.Options.OrderBy)
            {
                case ContentsOrder.Modified:
                    query = query.OrderByDescending<CommonPartRecord>(cr => cr.ModifiedUtc);
                    break;
                case ContentsOrder.Published:
                    query = query.OrderByDescending<CommonPartRecord>(cr => cr.PublishedUtc);
                    break;
                case ContentsOrder.Created:
                    query = query.OrderByDescending<CommonPartRecord>(cr => cr.CreatedUtc);
                    break;
            }

            //if (!String.IsNullOrWhiteSpace(model.Options.SelectedCulture))
            //{
            //    query = _cultureFilter.FilterCulture(query, model.Options.SelectedCulture);
            //}

            //if (model.Options.ContentsStatus == ContentsStatus.Owner)
            //{
            //    query = query.Where<CommonPartRecord>(cr => cr.OwnerId == Services.WorkContext.CurrentUser.Id);
            //}
            if (!string.IsNullOrEmpty(model.Code))
            {
                int id = 0;
                int.TryParse(model.Code, out id);
                query = query.Where<CommonPartRecord>(x => x.Id == id);
            }
            if (!string.IsNullOrWhiteSpace(model.PublishTime))
            {
                DateTime date = DateTime.Parse(model.PublishTime);
                query = query.Where<CommonPartRecord>(cr => cr.PublishedUtc >= date);
                //listExportDatas = listExportDatas.Where(x => DateTime.Parse(x.PublishTime) >= date);
            }

            var maxCount = query.Count();
            listExportDatas = query.Slice(0, maxCount).Select(x => GetListContentsModel(x));
            if (listExportDatas != null && listExportDatas.Any())
            {
                if (!string.IsNullOrEmpty(model.Keywords))
                {
                    listExportDatas = listExportDatas.Where(x => x.Keywords.Contains(model.Keywords));
                }
                if (!string.IsNullOrEmpty(model.Tags))
                {
                    //listExportDatas = listExportDatas.Where(x => x.Tags.Contains(model.Tags));
                }
            }

            if (listExportDatas != null && listExportDatas.Any()) maxCount = listExportDatas.Count();
            else maxCount = 0;

            model.Options.SelectedFilter = "News";// model.TypeName;
            //model.Options.FilterOptions = GetListableTypes(false)
            //    .Select(ctd => new KeyValuePair<string, string>(ctd.Name, ctd.DisplayName))
            //    .ToList().OrderBy(kvp => kvp.Value);

            //model.Options.Cultures = _cultureManager.ListCultures();

            var maxPagedCount = _siteService.GetSiteSettings().MaxPagedCount;
            if (maxPagedCount > 0 && pager.PageSize > maxPagedCount)
                pager.PageSize = maxPagedCount;
            var pagerShape = Shape.Pager(pager).TotalItemCount(maxPagedCount > 0 ? maxPagedCount : maxCount);
            var pageOfContentItems = query.Slice(pager.GetStartIndex(), pager.PageSize).ToList();

            var list = new List<ListContentsModel>();
            if (listExportDatas != null && listExportDatas.Any())
            {
                list = listExportDatas.Skip(pager.GetStartIndex()).Take(pager.PageSize).ToList();
            }

            var viewModel = Shape.ViewModel()
               .ContentItems(list)
               .Pager(pagerShape)
               .Options(model);

            //ViewBag.VeiwModel = model;

            return View(viewModel);
        }

        /// <summary>
        /// 使用ajax导出存在一个问题：生成文件地址在客户端访问不到，待解决
        /// </summary>
        /// <param name="__RequestVerificationToken"></param>
        /// <returns></returns>
        [ActionName("ExportPost")]
        [HttpPost]
        public JsonResult Export(string __RequestVerificationToken)
        {
            if (listCols == null)
            {
                listCols = new List<ExcelColumns>() {
                    new ExcelColumns("Id","ID"),
                    new ExcelColumns("Title","标题"),
                    new ExcelColumns("PublishTime","发布时间"),
                    new ExcelColumns("Keywords","关键词"),
                    new ExcelColumns("LinkUrl","地址"),
                };
            }

            if (listExportDatas == null || listExportDatas.Count() == 0)
            {
                return Json(new { status = 0, msg = "暂无数据" }, JsonRequestBehavior.AllowGet);
            }
            var fileKey = string.Format("{0}_{1}{2}", "资讯统计", DateTime.Now.ToString("yyyyMMddHHmmss"), new Random().Next(1000, 9999).ToString());
            try
            {
                System.Data.DataTable dt = Utils.ListToTable(listExportDatas.ToList(), listCols.Select(x => x.FieldName).ToArray());
                var file = ExcelHelper.instance.ExportFromTable(dt, listColumns: listCols, fileName: fileKey, isAjax: true);
                return Json(new { status = 1, msg = file }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "啊彩资讯导出失败");
            }

            return Json(new { status = 0, msg = "导出失败" }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost, ActionName("Export")]
        public ActionResult ExportPOST()
        {
            if (listExportDatas == null)
            {
                Response.Write("<script>alert('暂无数据，请刷新重新操作');</script>");
                return RedirectToAction("List", "admin", new { area = "Acai.NewsStatistics" });
            }
            if (listCols == null)
            {
                listCols = new List<ExcelColumns>() {
                    new ExcelColumns("Id","ID"),
                    new ExcelColumns("Title","标题"),
                    new ExcelColumns("PublishTime","发布时间"),
                    new ExcelColumns("Keywords","关键词"),
                    new ExcelColumns("LinkUrl","地址"),
                };
            }
            var fileKey = string.Format("{0}_{1}{2}", "资讯统计", DateTime.Now.ToString("yyyyMMddHHmmss"), new Random().Next(1000, 9999).ToString());
            var exportFilePath = "";
            try
            {
                System.Data.DataTable dt = Utils.ListToTable(listExportDatas.ToList(), listCols.Select(x => x.FieldName).ToArray());
                exportFilePath = ExcelHelper.instance.ExportFromTable(dt, listColumns: listCols, fileName: fileKey, isAjax: false);
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "啊彩资讯导出失败");
            }

            return File(exportFilePath, "application /vnd.ms-excel", fileKey + ".xlsx");
        }

        private IEnumerable<ContentTypeDefinition> GetListableTypes(bool andContainable)
        {
            return _contentDefinitionManager.ListTypeDefinitions().Where(ctd =>
                Services.Authorizer.Authorize(Permissions.EditContent, _contentManager.New(ctd.Name)) &&
                ctd.Settings.GetModel<ContentTypeSettings>().Listable &&
                (!andContainable || ctd.Parts.Any(p => p.PartDefinition.Name == "ContainablePart")));
        }

        private ListContentsModel GetListContentsModel(ContentItem m)
        {
            var tmp = new ListContentsModel()
            {
                Id = m.Id,
            };
            if (!string.IsNullOrEmpty(m.Record.Data))
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(m.Record.Data);

                var publishTime = doc.SelectSingleNode("Data/CommonPart").Attributes["PublishedUtc"].Value;
                tmp.PublishTime = !string.IsNullOrEmpty(publishTime) ? Convert.ToDateTime(publishTime).ToString("yyyy/MM/dd HH:mm:ss") : "-";
                tmp.Keywords = doc.SelectSingleNode("Data/MetaPart").Attributes["Keywords"].Value;
            }
            if (!string.IsNullOrEmpty(m.VersionRecord.Data))
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(m.VersionRecord.Data);
                tmp.Title = doc.SelectSingleNode("Data/TitlePart").Attributes["Title"].Value;
                tmp.LinkUrl = $"http://www.acaicp.com/{doc.SelectSingleNode("Data/AutoroutePart").Attributes["DisplayAlias"].Value}.aspx".ToLower();
            }

            return tmp;
        }
    }
}