using System.Web.Http;

namespace SaveToSharepointWeb
{
    public class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            config.MapHttpAttributeRoutes();
            config.Routes.MapHttpRoute(
                name:"DefaultAPI",
                routeTemplate:"api/{controller}/{id}",
                defaults:new {id=RouteParameter.Optional}
                );
        }
    }
}