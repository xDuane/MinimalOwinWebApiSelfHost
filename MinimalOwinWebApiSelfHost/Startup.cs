using System;

// Add the following usings:
using Owin;
using System.Web;
using System.Web.Http;
using MinimalOwinWebApiSelfHost.Models;
using MinimalOwinWebApiSelfHost.OAuthServerProvider;
using Microsoft.Owin.Security.OAuth;
using Microsoft.Owin;
using Microsoft.Owin.FileSystems;
using Microsoft.Owin.StaticFiles;

using System.Runtime.Caching;

namespace MinimalOwinWebApiSelfHost
{
    public class Startup
    {
        // This method is required by Katana:
        public void Configuration(IAppBuilder app)
        {
            var cache = MemoryCache.Default;
            //TODO: Config
            cache["_cacheDays"] = 30;

            ConfigureAuth(app);

            var webApiConfiguration = ConfigureWebApi();
            app.UseWebApi(webApiConfiguration);

            app.MapSignalR();

            var physicalFileSystem = new PhysicalFileSystem(@".\Web");
            var options = new FileServerOptions
            {
                EnableDefaultFiles = true,
                FileSystem = physicalFileSystem
            };
            options.StaticFileOptions.FileSystem = physicalFileSystem;
            options.StaticFileOptions.ServeUnknownFileTypes = true;
            options.DefaultFilesOptions.DefaultFileNames = new[] { "home.html" };
            app.UseFileServer(options);

        }

        private void ConfigureAuth(IAppBuilder app)
        {
            var OAuthOptions = new OAuthAuthorizationServerOptions
            {
                TokenEndpointPath = new PathString("/Token"),
                Provider = new ApplicationOAuthServerProvider(),
                AccessTokenExpireTimeSpan = TimeSpan.FromDays(14),

                // Only do this for demo!!
                AllowInsecureHttp = true
            };
            app.UseOAuthAuthorizationServer(OAuthOptions);
            app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions());
        }
        
        private HttpConfiguration ConfigureWebApi()
        {
            var config = new HttpConfiguration();
            config.Routes.MapHttpRoute(
                "DefaultApi",
                "api/{controller}/{action}/{zip}", null
                //new { id = RouteParameter.Optional }
                );
            return config;
        }
    }
}
