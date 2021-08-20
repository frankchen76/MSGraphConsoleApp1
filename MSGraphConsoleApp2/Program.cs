using Azure.Identity;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSGraphConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            Test().GetAwaiter().GetResult();
        }
        static async Task Test()
        {
            // Please update to your client id and tenant id
            string clientId = "3fd52fc6-7f18-459b-b123-b5f7bce5bbf6";
            string tenantId = "8a5ee357-7de0-4836-ab20-9173b12cdce9";
            string[] scopes = { "Calendars.ReadWrite", 
                "Directory.AccessAsUser.All", 
                "Group.Read.All",
                "Mail.Read", 
                "Mail.Send", 
                "profile", 
                "openid", 
                "email" };

            var loggerFactory = LoggerFactory.Create(builder =>
            {
                builder
                    .AddFilter("Microsoft", Microsoft.Extensions.Logging.LogLevel.Information)
                    .AddFilter("System", Microsoft.Extensions.Logging.LogLevel.Information)
                    .AddFilter("NonHostConsoleApp.Program", Microsoft.Extensions.Logging.LogLevel.Information)
                    .AddConsole();
            });
            ILogger< LoggingHandler> logger = loggerFactory.CreateLogger<LoggingHandler>();

            IPublicClientApplication publicClientApplication = PublicClientApplicationBuilder
                            .Create(clientId)
                            .WithTenantId(tenantId) 
                            .WithRedirectUri("http://localhost")
                            .Build();

            InteractiveAuthenticationProvider authenticationProvider = new InteractiveAuthenticationProvider(publicClientApplication, scopes);

            // get the default list of handlers and add the logging handler to the list
            var handlers = GraphClientFactory.CreateDefaultHandlers(authenticationProvider);
            handlers.Add(new LoggingHandler(logger));

            // create the GraphServiceClient with logging support
            var httpClient = GraphClientFactory.Create(handlers);
            GraphServiceClient graphClient = new GraphServiceClient(httpClient);

            //User me = await graphClient.Me.Request().GetAsync();
            //Console.WriteLine($"{me.ToString()}");

            await CreateEventWithExtendedPropertiesWithClientData(graphClient);
            Console.WriteLine("Done");


        }
        static public async Task CreateEventWithExtendedPropertiesWithClientData(GraphServiceClient graphClient)
        {
            var subject = "Interview - Test Test";
            var startDateTime = new DateTime(2021, 08, 21, 11, 0, 0);
            var endDateTime = startDateTime.AddMinutes(45);
            var timezone = "Eastern Standard Time";
            var location = "";
            var content = "<br />Date: Aug 20, 2021<br />Time: 09:00 AM ET<br />Duration: 45 min<br />Location: <br /><br />To view information about the applicant click:<br />https://viwavetest.viglobalcloud.com/Ashley/WG_Sacks/viRecruitAplInfo/login.aspx?CompanyID=2005061717&REID=5&ApplicantID=7728A92789DF4FA6AD15&ApplicationID=11E5CCA10E6A48B09536 /><br />To view the applicant interview schedule click:<br />https://viwavetest.viglobalcloud.com/Ashley/WG_Sacks/viRecruitAplInfo/ReApplicantInterviewPrint.aspx?CompanyID=2005061717&REID=5&ApplicantID=7728A92789DF4FA6AD15&ApplicationID=11E5CCA10E6A48B09536 /><br />";
            var viDesktopMeetingId = "117d94d5-05d1-4016-beed-4c519358d174";



            var @event = new Event
            {
                Subject = subject,
                Start = new DateTimeTimeZone() { DateTime = startDateTime.ToString("s"), TimeZone = timezone },
                End = new DateTimeTimeZone() { DateTime = endDateTime.ToString("s"), TimeZone = timezone },
                Location = new Location() { DisplayName = location },
                Body = new ItemBody() { ContentType = BodyType.Text, Content = content },
                Sensitivity = Sensitivity.Normal,
                ResponseRequested = true,
                ReminderMinutesBeforeStart = 15
            };



            if (!string.IsNullOrEmpty(viDesktopMeetingId))
            {
                @event.SingleValueExtendedProperties = new EventSingleValueExtendedPropertiesCollectionPage() {
                    new SingleValueLegacyExtendedProperty() { Id = $"String {{{ Guid.NewGuid().ToString("D")}}} Name ViDesktopMeetingId", Value = viDesktopMeetingId }
                };
            }
            //var createdEvent = await graphService.CreateEventAsync(@event);
            var createdEvent = await graphClient.Me.Calendar.Events.Request().AddAsync(@event).ConfigureAwait(false);

            Console.WriteLine("Done");
        }
    }
}
