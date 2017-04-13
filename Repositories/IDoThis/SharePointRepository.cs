using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sierra.Azure.CommonDemoAPI.Models.IDoThis;
using Sierra.SharePoint.Library.CSOM;
using Sierra.NET.Core;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace Sierra.Azure.CommonDemoAPI.Repositories.IDoThis
{
    public class SharePointRepository : IDoThisRepository
    {
        private string _configuration;
        public SharePointRepository(string configuration)
        {
            _configuration = configuration;
        }

        public UserProfile GetUserProfile(string userId)
        {
            try
            {
                SimpleTextLogger logger = new SimpleTextLogger();

                //cannot use these lines as they cause error -- need to use routines below
                var credentials = System.Net.CredentialCache.DefaultCredentials;
                SPClientUtility context = new SPClientUtility(logger, credentials);


                var lists = context.GetAllListTitles("https://sierrasystemsgroup.sharepoint.com/sites/siza");

                return new UserProfile { Id = userId, Name = "user from SharePoint repository" };
            }
            catch(Exception ex)
            {
                return new UserProfile { Id = userId, Name = ex.Message };
            }
        }

        //private async Task<string> GetAccessToken()
        //{
        //    string clientId = "todo";// ConfigurationManager.AppSettings["ida:ClientId"];
        //    string appKey = "todo";// ConfigurationManager.AppSettings["ida:AppKey"];
        //    string aadInstance = "todo";// ConfigurationManager.AppSettings["ida:AADInstance"];
        //    string domain = "todo";// ConfigurationManager.AppSettings["ida:Domain"];
        //    string resource = "todo";// ConfigurationManager.AppSettings["ida:Resource"];

        //    AuthenticationResult result = null;

        //    ClientCredential clientCred = new ClientCredential(clientId, appKey);
        //    string authHeader = HttpContext.Current.Request.Headers["Authorization"];
        //    string userAccessToken = authHeader.Substring(authHeader.LastIndexOf(' ')).Trim();
        //    UserAssertion userAssertion = new UserAssertion(userAccessToken);
        //    string authority = aadInstance + domain;
        //    AuthenticationContext authContext = new AuthenticationContext(authority);

        //    //result = await authContext.AcquireTokenAsync(resource, clientCred); // auth without user assertion (fails, app only not allowed)

        //    result = await authContext.AcquireTokenAsync(resource, clientCred, userAssertion); 
        //    return result.AccessToken;
        //}


        //public ClientContext GetAzureADAccessTokenAuthenticatedContext(String siteUrl, String accessToken)
        // { 
        //     var clientContext = new ClientContext(siteUrl); 
 
 
        //     clientContext.ExecutingWebRequest += (sender, args) => 
        //     { 
        //         args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken; 
        //     }; 
 
 
        //     return clientContext; 
        // } 
 
 


    }
}
