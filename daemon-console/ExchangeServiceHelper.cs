using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Threading.Tasks;
using static daemon_console.Program;
using System.Collections.Generic;

namespace daemon_console
{

    public class AzureAdConfig
    {
        public string Instance { get; set; }
        public string TenantId { get; set; }
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
        public string Authority { get; set; }

        
    }

    public class ExchangeServiceHelper
    {
        private readonly AzureAdConfig _config;

        public ExchangeServiceHelper(IConfiguration configuration)
        {
            _config = configuration.GetSection("AzureAd").Get<AzureAdConfig>();
        }

        public async Task<ExchangeService> GetExchangeServiceAsync()
        {
            // Initialize MSAL
            IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(_config.ClientId)
                .WithClientSecret(_config.ClientSecret)
                .WithAuthority(new Uri(_config.Instance))
                .Build();

            // Acquire the token
            AuthenticationResult result = await app.AcquireTokenForClient(new[] { "https://outlook.office365.com/.default" }).ExecuteAsync();

            // Initialize the ExchangeService with OAuth
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
            {
                Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx")
            };
            service.Credentials = new OAuthCredentials(result.AccessToken);

            // Set the impersonation property
            // "albertly_dev@albertly.onmicrosoft.com"
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "pavel@albertly.onmicrosoft.com");

            return service;
        }

        public async Task<List<EmailMessage>> ReadEmailsAsync()
        {
            ExchangeService service = await GetExchangeServiceAsync();


            // Define the search filters
            SearchFilter unreadFilter = new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false);
            SearchFilter dateFilter = new SearchFilter.IsGreaterThanOrEqualTo(EmailMessageSchema.DateTimeReceived, DateTime.Now.AddDays(-1));

            // Combine the filters
            SearchFilter.SearchFilterCollection searchFilterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.And, unreadFilter, dateFilter);


            // Define the search filter and view
            ItemView view = new ItemView(10); // Fetch the top 10 emails
            
            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, searchFilterCollection, view);

            List<EmailMessage> emails = new List<EmailMessage>();
            foreach (Item item in findResults.Items)
            {
                if (item is EmailMessage email)
                {
                    email.Load(); // Load the email details
                    emails.Add(email);
                }
            }

            return emails;
        }
    }
}