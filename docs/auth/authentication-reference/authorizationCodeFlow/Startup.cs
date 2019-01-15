using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace authorizationCodeFlow
{
    public class Startup
    {
        private readonly string TenantName = "TXu4dow.onmicrosoft.com";
        private readonly string ClientId = "b93f9575-17ec-44d2-8db0-befa0cab7888";
        private readonly string ClientSecret = "<clientsecret>";
        // This method gets called by the runtime. Use this method to add services to the container.
        // For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=398940
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddAuthentication(options =>
                    {
                        options.DefaultScheme = CookieAuthenticationDefaults.AuthenticationScheme;
                        options.DefaultChallengeScheme = OpenIdConnectDefaults.AuthenticationScheme;
                    })
                    .AddOpenIdConnect(options =>
                    {
                        /* this section has information used to obtain an authorization code from Azure */
                        options.Authority = "https://login.microsoftonline.com/" + this.TenantName;
                        // application id in Azure
                        options.ClientId = this.ClientId;
                        // the redirect_uri, must match one on Azure
                        options.CallbackPath = "/security/signin-callback";
                        // using the authorization code flow
                        options.ResponseType = OpenIdConnectResponseType.Code;
                        /* end section */

                        /* this section has additional information used to obtain an ID/access token from Azure */
                        // portal >> registered app >> keys
                        options.ClientSecret = this.ClientSecret;
                        options.SaveTokens = true;
                        options.Events = new OpenIdConnectEvents
                        {
                            OnAuthorizationCodeReceived = async (ContextBoundObject) =>
                            {
                                ClientCredential credentials = new ClientCredential(this.ClientId, this.ClientSecret);

                                var authContext = new AuthenticationContext(ContextBoundObject.Options.Authority);
                                // last variable (resource) is the id of the client app, which receives the token
                                var result = await authContext.AcquireTokenByAuthorizationCodeAsync(
                                    ContextBoundObject.ProtocolMessage.Code,
                                    new Uri(ContextBoundObject.Properties.Items[OpenIdConnectDefaults.RedirectUriForCodePropertiesKey]), credentials, this.ClientId);

                                // tell middleware that we already handled the code redemption
                                ContextBoundObject.HandleCodeRedemption(result.AccessToken, result.IdToken);
                            }
                        };
                        /* end section */
                    })
                    .AddCookie();

            services.AddMvc();
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            app.UseAuthentication();

            app.UseMvc(routes =>
            {
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
