using System.Collections.Generic;
using System.Web.Mvc;
using System.Web.Routing;
using Orchard.Mvc.Routes;

namespace Acai.NewsStatistics
{
    public class Routes : IRouteProvider
    {
        public void GetRoutes(ICollection<RouteDescriptor> routes)
        {
            foreach (var routeDescriptor in GetRoutes())
                routes.Add(routeDescriptor);
        }

        public IEnumerable<RouteDescriptor> GetRoutes()
        {
            return new[] {
                new RouteDescriptor {
                    Route = new Route(
                        "Acai.NewsStatistics", // this is the name of the page url
                        new RouteValueDictionary {
                            {"area", "Acai.NewsStatistics"}, // this is the name of your module
                            {"controller", "Home"},
                            {"action", "Index"}
                        },
                        new RouteValueDictionary(),
                        new RouteValueDictionary {
                            {"area", "Acai.NewsStatistics"} // this is the name of your module
                        },
                        new MvcRouteHandler())
                },
                 new RouteDescriptor {
                    Route = new Route(
                        "Acai.NewsStatistics",
                        new RouteValueDictionary {
                            {"area", "Acai.NewsStatistics"},
                            {"controller", "Admin"},
                            {"action", "ExportPost"}
                        },
                        new RouteValueDictionary(),
                        new RouteValueDictionary {
                            {"area", "Acai.NewsStatistics"}
                        },
                        new MvcRouteHandler())
                }
            };
        }
    }
}
