using GraphsPOC.Authentication;
using GraphTutorial;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using Microsoft.Identity.Client;
using System.ServiceModel;
using Microsoft.Graph.Auth;

namespace GraphsPOC
{
    class Program
    {
        static IConfigurationRoot LoadAppSettings()
        {
            // dotnet user-secrets set <appId/scopes> <value>
            // to set settings
            var appConfig = new ConfigurationBuilder()
                .AddUserSecrets<Program>()
                .Build();

            // Check for required settings
            if (string.IsNullOrEmpty(appConfig["appId"]) ||
                string.IsNullOrEmpty(appConfig["scopes"]))
            {
                return null;
            }


            return appConfig;
        }

        static string FormatDateTimeTimeZone(
            Microsoft.Graph.DateTimeTimeZone value,
            string dateTimeFormat)
        {
            // Parse the date/time string from Graph into a DateTime
            var dateTime = DateTime.Parse(value.DateTime);

            return dateTime.ToString(dateTimeFormat);
        }

        static void Main(string[] args)
        {
            var appConfig = LoadAppSettings();

            if (appConfig == null)
            {
                Console.WriteLine("Missing or invalid appsettings.json...exiting");
                return;
            }

            var clientSecret = "CLIENT-SECRET";

            var appId = appConfig["appId"];
            var scopesString = appConfig["scopes"];
            var scopes = scopesString.Split(';');

            // Initialize the auth provider with values from appsettings.json
            IConfidentialClientApplication app;
            app = ConfidentialClientApplicationBuilder
                .Create(appConfig["appId"])
                .WithClientSecret(clientSecret)
                .WithTenantId("TENANT-ID")
                .Build();

            ClientCredentialProvider authProvider = new ClientCredentialProvider(app);
            // Request a token to sign in the user
            //var accessToken = authProvider.GetAccessToken().Result;

            //Console.WriteLine($"Access token: {accessToken}\n");

            // Initialize Graph client
            GraphHelper.Initialize(authProvider);

            // Get signed in user
            //var user = GraphHelper.GetMeAsync().Result;
            //Console.WriteLine($"Welcome {user.DisplayName}!\n");

            GraphServiceClient graphClient = new GraphServiceClient(authProvider);

            var @event = new Event
            {
                Subject = "Let's go for lunch",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = "Does next month work for you?"
                },
                Start = new DateTimeTimeZone
                {
                    DateTime = "2021-07-15T12:00:00",
                    TimeZone = "Pacific Standard Time"
                },
                End = new DateTimeTimeZone
                {
                    DateTime = "2021-07-15T14:00:00",
                    TimeZone = "Pacific Standard Time"
                },
                Location = new Location
                {
                    DisplayName = "Test Location"
                },
                Attendees = new List<Attendee>()
                {
                    new Attendee
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = "EMAIL ADDRESS",
                            Name = "EMAIL"
                        },
                        Type = AttendeeType.Required
                    }
                },
                    IsOnlineMeeting = true,
                    OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness
            };


            var e = graphClient.Users["EXCHANGE-ACCOUNT"].Events.Request().AddAsync(@event).Result;

            Console.WriteLine("Event should now be scheduled");
        }
    }
}
