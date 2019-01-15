using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.Extensions.Options;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.DataProtection;
using System.Security.Claims;
using System.IO;
using System.Text;

namespace authentication
{
    public class Startup
    {
        // This method gets called by the runtime. Use this method to add services to the container.
        // For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=398940
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddDataProtection().PersistKeysToFileSystem(new DirectoryInfo(@"C:\Git\stuffstuff\"));

            services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
                .AddCookie(options => {
                    options.LoginPath = "/Account/Index/";
                    options.DataProtectionProvider = DataProtectionProvider.Create(new DirectoryInfo(@"C:\Git\stuffstuff\"));
                    options.Events = new CookieAuthenticationEvents
                    {
                        OnValidatePrincipal = async (ContextBoundObject) =>
                        {
                            //GetClaimFromCookie(ContextBoundObject.HttpContext, ".AspNetCore.Cookies", "");
                        }
                    };
                });

            services.AddMvc();
        }

        private IEnumerable<Claim> GetClaimFromCookie(HttpContext httpContext, string cookieName, string cookieSchema)
        {
            // Get the encrypted cookie value
            var opt = httpContext.RequestServices.GetRequiredService<IOptionsMonitor<CookieAuthenticationOptions>>();
            var cookie = opt.CurrentValue.CookieManager.GetRequestCookie(httpContext, cookieName);

            // Decrypt if found
            if (!string.IsNullOrEmpty(cookie))
            {
                var provider = DataProtectionProvider.Create(new DirectoryInfo(@"C:\Git\stuffstuff\"));
                var dataProtector = provider.CreateProtector(typeof(CookieAuthenticationDefaults).FullName, "Cookies", "v2");

                UTF8Encoding specialUtf8Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);
                byte[] protectedBytes = Base64UrlTextEncoder.Decode(cookie);
                byte[] plainBytes = dataProtector.Unprotect(protectedBytes);
                string plainText = specialUtf8Encoding.GetString(plainBytes);

                var ticketDataFormat = new TicketDataFormat(dataProtector);
                var ticket = ticketDataFormat.Unprotect(cookie);
                return ticket.Principal.Claims;
            }
            return null;
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        {
            app.UseAuthentication();

            app.UseMvc(routes => {
                routes.MapRoute(
                    name: "default",
                    template: "{controller=Home}/{action=Index}/{id?}");
            });

            app.Run(async (context) =>
            {
                await context.Response.WriteAsync("Hello World!");
            });
        }
    }
}
