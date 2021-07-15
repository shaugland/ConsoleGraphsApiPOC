using GraphsPOC.Authentication;
using GraphTutorial;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System;
using System.Collections.Generic;

namespace GraphsPOC
{
    class Program
    {
        static IConfigurationRoot LoadAppSettings()
        {
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

            var appId = appConfig["appId"];
            var scopesString = appConfig["scopes"];
            var scopes = scopesString.Split(';');

            // Initialize the auth provider with values from appsettings.json
            var authProvider = new DeviceCodeAuthProvider(appId, scopes);

            // Request a token to sign in the user
            var accessToken = authProvider.GetAccessToken().Result;

            Console.WriteLine($"Access token: {accessToken}\n");

            // Initialize Graph client
            GraphHelper.Initialize(authProvider);

            // Get signed in user
            var user = GraphHelper.GetMeAsync().Result;
            Console.WriteLine($"Welcome {user.DisplayName}!\n");

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
                            Address = "shauglan_stu@kent.edu",
                            Name = "Seth Haugland"
                        },
                        Type = AttendeeType.Required
                    }
                },
                    IsOnlineMeeting = true,
                    OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness
            };

            var something = graphClient.Me.Calendar.Events
                .Request()
                .AddAsync(@event).Result;

            Console.WriteLine("Event should now be scheduled");
        }
    }
}
