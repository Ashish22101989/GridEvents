using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;
using System.Net;
using System.ServiceModel.Description;
using Microsoft.Xrm.Tooling.Connector;
using System.Net.Http;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Net.Http.Headers;

namespace Appointment
{
    class Program
    {
        static void Main(string[] args)
        {
            // Dynamics CRM Online Instance URL
            string resource = "https://orgf86c4a8e.crm8.dynamics.com";

            // client id and client secret of the application
            ClientCredential clientCrendential = new ClientCredential("fa633876-9a9d-48e9-aba8-0e30e828bbb7", "390e7df5-be85-4cb7-b09d-a7f98a296596");
            

            // Authenticate the registered application with Azure Active Directory.
            AuthenticationContext authContext =
            new AuthenticationContext("https://login.microsoftonline.com/e4009f96-f22f-42d4-959f-e4862e96d7e9/oauth2/token");

            Task<AuthenticationResult> authResult = authContext.AcquireTokenAsync(resource,clientCrendential);
            var accessToken = authResult.Result.AccessToken;

            // use HttpClient to call the Web API
            HttpClient httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
            httpClient.DefaultRequestHeaders.Add("OData-Version", "4.0");
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            httpClient.BaseAddress = new Uri("https://orgf86c4a8e.crm8.dynamics.com/api/data/v9.0/");

            var response = httpClient.GetAsync("WhoAmI").Result;
            if (response.IsSuccessStatusCode)
            {
                var userDetails = response.Content.ReadAsStringAsync().Result;
            }

        }
    }
}
