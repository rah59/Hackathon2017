using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.OpenIdConnect;
using Microsoft.Office365.OutlookServices;

using dotnet_tutorial.TokenStorage;
using System.Configuration;
using WebApplication1.Models;
using System.Net;
using System.IO;
using System.Net.Http;
using Newtonsoft.Json;
using RestSharp;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            if (Request.IsAuthenticated)
            {
                string userName = ClaimsPrincipal.Current.FindFirst("name").Value;
                string userId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
                if (string.IsNullOrEmpty(userName) || string.IsNullOrEmpty(userId))
                {
                    // Invalid principal, sign out
                    return RedirectToAction("SignOut");
                }

                // Since we cache tokens in the session, if the server restarts
                // but the browser still has a cached cookie, we may be
                // authenticated but not have a valid token cache. Check for this
                // and force signout.
                SessionTokenCache tokenCache = new SessionTokenCache(userId, HttpContext);
                if (tokenCache.Count <= 0)
                {
                    // Cache is empty, sign out
                    return RedirectToAction("SignOut");
                }

                ViewBag.UserName = userName;
            }
            return View();
        }


        public class CalHint
        {
            public int eventID { get; set; }
            public int origEmailID { get; set; }
            public int hostUserID { get; set; }
            public string date { get; set; }
            public string time { get; set; }
            public string meetWithName { get; set; }
            public string meetWithAddress { get; set; }
            public string subject { get; set; }
            public string eventNotes { get; set; }
        }

        public class RootObject
        {
            public int emailID { get; set; }
            public string sender { get; set; }
            public int senderID { get; set; }
            public string recipient { get; set; }
            public int recipientID { get; set; }
            public string subject { get; set; }
            public long dateSent { get; set; }
            public string messageBody { get; set; }
            public string processedMessage { get; set; }
            public List<CalHint> calHints { get; set; }
        }

        public async Task<ActionResult> Detail(string subject)
        {
            string token = await GetAccessToken();
            if (string.IsNullOrEmpty(token))
            {
                // If there's no token in the session, redirect to Home
                return Redirect("/");
            }

            string userEmail = await GetUserEmail();

            OutlookServicesClient client =
                new OutlookServicesClient(new Uri("https://outlook.office.com/api/v2.0"), GetAccessToken);

            client.Context.SendingRequest2 += new EventHandler<Microsoft.OData.Client.SendingRequest2EventArgs>(
                (sender, e) => InsertXAnchorMailboxHeader(sender, e, userEmail));

            try
            {
                //Grabs all messages from inbox
                var mailResults = await client.Me.Messages
                                    .OrderByDescending(m => m.ReceivedDateTime)
                                    .Take(5)
                                    .Select(m => new Models.DisplayMessage(m.Subject, m.ReceivedDateTime, m.Sender, m.Body, m.From.EmailAddress))
                                    .ExecuteAsync();
                //Finds the email that matches the subject of the selected email
                var CurrentPage = mailResults.CurrentPage.Where(i => i.Subject == subject).First();

                //Build the post object with parameters
                System.Text.ASCIIEncoding encoding = new System.Text.ASCIIEncoding();
                   
                // Post the information to the NLP end point and retrieve the email id
                  //string postData = 
                  //"{ \"sender\": \"" + CurrentPage.From + "\"," + 
                  //"\"recipient\": \"" + CurrentPage.Recipient + "\"," +
                  //"\"subject\" : \"" + CurrentPage.Subject + "\"," +
                  //"\"messageBody\" : \"" + CurrentPage.Body + "\"}";

                //byte[] data = encoding.GetBytes(postData);

                ////HTTP request object creation
                //HttpWebRequest myRequest =
                //(HttpWebRequest)WebRequest.Create("http://70.94.39.41:8080/email");
                //myRequest.Method = "POST";
                //myRequest.ContentType = "application/x-www-form-urlencoded";
                //myRequest.ContentLength = data.Length;
                //Stream newStream = myRequest.GetRequestStream();
                ////Data posted to URL
                //newStream.Write(data, 0, data.Length);
                //newStream.Close();


                // Then use the email id to retrieve the modified email with links to events detected in the email.
                // For the purposes of the demo, we will use an email that has been converted already

                var gclient = new RestClient("http://70.94.39.41:8080/email?id=20");
                var request = new RestRequest(Method.GET);
                request.AddHeader("postman-token", "3ee3f84c-6045-c1f6-4c52-7afe5822c179");
                request.AddHeader("cache-control", "no-cache");
                request.AddHeader("content-type", "application/json");
                request.AddParameter("application/json", "{ \"sender\": \"clthf@mail.umkc.edu\",\n  \"recipient\": \"rah59@mail.umkc.edu\",\n  \"subject\" : \"Would like to check out the car\",\n  \"messageBody\" : \"Let's get together a week from tomorrow and see if we can figure it out.\"\n}", ParameterType.RequestBody);
                IRestResponse response = gclient.Execute(request);

                string responseJson = response.Content;

                RootObject respObj = JsonConvert.DeserializeObject<RootObject>(responseJson);

                CurrentPage.Body = respObj.processedMessage;
                

                return View(CurrentPage);
            }
            catch (MsalException ex)
            {
                return RedirectToAction("Error", "Home", new { message = "ERROR retrieving messages", debug = ex.Message });
            }
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult Error(string message, string debug)
        {
            ViewBag.Message = message;
            ViewBag.Debug = debug;
            return View("Error");
        }

        public void SignIn()
        {
            if (!Request.IsAuthenticated)
            {
                // Signal OWIN to send an authorization request to Azure
                HttpContext.GetOwinContext().Authentication.Challenge(
                    new AuthenticationProperties { RedirectUri = "/" },
                    OpenIdConnectAuthenticationDefaults.AuthenticationType);
            }
        }

        public void SignOut()
        {
            if (Request.IsAuthenticated)
            {
                string userId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;

                if (!string.IsNullOrEmpty(userId))
                {
                    string appId = ConfigurationManager.AppSettings["ida:AppId"];
                    // Get the user's token cache and clear it
                    SessionTokenCache tokenCache = new SessionTokenCache(userId, HttpContext);
                    tokenCache.Clear(appId);
                }
            }
            // Send an OpenID Connect sign-out request. 
            HttpContext.GetOwinContext().Authentication.SignOut(
                CookieAuthenticationDefaults.AuthenticationType);
            Response.Redirect("/");
        }

        public async Task<string> GetAccessToken()
        {
            string accessToken = null;

            // Load the app config from web.config
            string appId = ConfigurationManager.AppSettings["ida:AppId"];
            string appPassword = ConfigurationManager.AppSettings["ida:AppPassword"];
            string redirectUri = ConfigurationManager.AppSettings["ida:RedirectUri"];
            string[] scopes = ConfigurationManager.AppSettings["ida:AppScopes"]
                .Replace(' ', ',').Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            // Get the current user's ID
            string userId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;

            if (!string.IsNullOrEmpty(userId))
            {
                // Get the user's token cache
                SessionTokenCache tokenCache = new SessionTokenCache(userId, HttpContext);

                ConfidentialClientApplication cca = new ConfidentialClientApplication(
                    appId, redirectUri, new ClientCredential(appPassword), tokenCache);

                // Call AcquireTokenSilentAsync, which will return the cached
                // access token if it has not expired. If it has expired, it will
                // handle using the refresh token to get a new one.
                AuthenticationResult result = await cca.AcquireTokenSilentAsync(scopes);

                accessToken = result.Token;
            }

            return accessToken;
        }

        public async Task<ActionResult> Inbox()
        {
            string token = await GetAccessToken();
            if (string.IsNullOrEmpty(token))
            {
                // If there's no token in the session, redirect to Home
                return Redirect("/");
            }

            string userEmail = await GetUserEmail();

            OutlookServicesClient client =
                new OutlookServicesClient(new Uri("https://outlook.office.com/api/v2.0"), GetAccessToken);

            client.Context.SendingRequest2 += new EventHandler<Microsoft.OData.Client.SendingRequest2EventArgs>(
                (sender, e) => InsertXAnchorMailboxHeader(sender, e, userEmail));

            try
            {
                var mailResults = await client.Me.Messages
                                    .OrderByDescending(m => m.ReceivedDateTime)
                                    .Take(5)
                                    .Select(m => new Models.DisplayMessage(m.Subject, m.ReceivedDateTime, m.Sender, m.Body, m.From.EmailAddress))
                                    .ExecuteAsync();

                return View(mailResults.CurrentPage);
            }
            catch (MsalException ex)
            {
                return RedirectToAction("Error", "Home", new { message = "ERROR retrieving messages", debug = ex.Message });
            }
        }

        public async Task<string> GetUserEmail()
        {
            OutlookServicesClient client =
                new OutlookServicesClient(new Uri("https://outlook.office.com/api/v2.0"), GetAccessToken);

            try
            {
                var userDetail = await client.Me.ExecuteAsync();
                return userDetail.EmailAddress;
            }
            catch (MsalException ex)
            {
                return string.Format("#ERROR#: Could not get user's email address. {0}", ex.Message);
            }
        }

        private void InsertXAnchorMailboxHeader(object sender, Microsoft.OData.Client.SendingRequest2EventArgs e, string email)
        {
            e.RequestMessage.SetHeader("X-AnchorMailbox", email);
        }

       
    }
}