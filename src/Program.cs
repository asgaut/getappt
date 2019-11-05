using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using CommandLine;

// Background information:
// https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/get-started-with-ews-client-applications
// https://stackoverflow.com/questions/43759529/get-appointments-from-coworker-via-ews-only-with-free-busy-time-subject-loc/43759990#43759990

// This requires nuget packages:
// Install-Package Exchange.WebServices.Managed.Api -Version 2.2.1.2
// Install-Package Newtonsoft.Json -Version 12.0.2
// Install-Package CommandLineParser -Version 2.6.0

namespace GetAppt
{
    class Program
    {
        class Data
        {
            public DateTime querytime;
            public Dictionary<string, System.Collections.ObjectModel.Collection<CalendarEvent>> appointments;
            public Dictionary<int, string> freebusystatusmap;
        }

        public class Options
        {
            [CommandLine.Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
            public bool Verbose { get; set; }
            [CommandLine.Option('t', "trace", Required = false, HelpText = "Enable tracing output.")]
            public bool Tracing { get; set; }
        }

        static string GetCalItems(Options cmdline, string[] users, string serverUrl)
        {
            var service = new ExchangeService(ExchangeVersion.Exchange2016)
            {
                Credentials = new WebCredentials() // use default network credentials
            };
            service.TraceEnabled = cmdline.Tracing;
            service.TraceFlags = TraceFlags.All;

            if (users.Length == 0)
            {
                throw new ApplicationException("List of users is empty");
            }

            if (cmdline.Verbose)
            {
                Console.WriteLine("Using '{0}' as server URL", serverUrl);
            }
            service.Url = new Uri(serverUrl);

            const int NUM_DAYS = 5;

            // Create a collection of attendees. 
            List<AttendeeInfo> attendees = new List<AttendeeInfo>();
            foreach (var u in users)
            {
                attendees.Add(new AttendeeInfo()
                {
                    SmtpAddress = u,
                    AttendeeType = MeetingAttendeeType.Required
                });
            }

            // Specify options to request free/busy information and suggested meeting times.
            AvailabilityOptions availabilityOptions = new AvailabilityOptions
            {
                GoodSuggestionThreshold = 49,
                MaximumNonWorkHoursSuggestionsPerDay = 0,
                MaximumSuggestionsPerDay = 2,
                // Note that 60 minutes is the default value for MeetingDuration, but setting it explicitly for demonstration purposes.
                MeetingDuration = 60,
                MinimumSuggestionQuality = SuggestionQuality.Good,
                DetailedSuggestionsWindow = new TimeWindow(DateTime.Now.AddDays(0), DateTime.Now.AddDays(NUM_DAYS)),
                RequestedFreeBusyView = FreeBusyViewType.Detailed
            };

            // Return free/busy information and a set of suggested meeting times. 
            // This method results in a GetUserAvailabilityRequest call to EWS.
            GetUserAvailabilityResults results = service.GetUserAvailability(attendees,
                                                                             availabilityOptions.DetailedSuggestionsWindow,
                                                                             AvailabilityData.FreeBusyAndSuggestions,
                                                                             availabilityOptions);

            Data data = new Data
            {
                querytime = DateTime.Now,
                appointments = new Dictionary<string, System.Collections.ObjectModel.Collection<CalendarEvent>>(),
                freebusystatusmap = new Dictionary<int, string>()
                {
                    { 0, "Free" }, //     The time slot associated with the appointment appears as free.
                    { 1, "Tentative" }, //     The time slot associated with the appointment appears as tentative.
                    { 2, "Busy" }, //     The time slot associated with the appointment appears as busy.
                    { 3, "OOF" }, //     The time slot associated with the appointment appears as Out of Office.
                    { 4, "WorkingElsewhere" }, //     The time slot associated with the appointment appears as working else where.
                    { 5, "NoData" } //     No free/busy status is associated with the appointment.
                }
            };
            for (int i = 0; i < results.AttendeesAvailability.Count; i++)
            {
                data.appointments[users[i]] = results.AttendeesAvailability[i].CalendarEvents;
            }
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);
            if (cmdline.Verbose)
            {
                Console.WriteLine(json);
            }
            return json;
        }

        static void HttpPut(Options cmdline, string json, string uri)
        {
            var req = System.Net.WebRequest.Create(uri);
            var data = Encoding.UTF8.GetBytes(json);
            req.Method = "PUT";
            req.Timeout = 10000;
            req.ContentType = "application/json";
            req.ContentLength = data.Length;

            using (System.IO.Stream sendStream = req.GetRequestStream())
            {
                if (cmdline.Verbose)
                {
                    Console.WriteLine("Performing http put to '{0}'", uri);
                }
                sendStream.Write(data, 0, data.Length);
                sendStream.Close();
                if (cmdline.Verbose)
                {
                    var res = req.GetResponse();
                    var buf = new byte[res.ContentLength];
                    using (var rxStream = res.GetResponseStream())
                    {
                        rxStream.Read(buf, 0, buf.Length);
                        var r = res as System.Net.HttpWebResponse;
                        Console.WriteLine("Response code: {0}, data:\r\n{1}", r.StatusCode, Encoding.UTF8.GetString(buf));
                    }
                }
            }
        }

        static void Main(string[] args)
        {
            var users = System.Configuration.ConfigurationManager.AppSettings["Users"].Split(';');
            var serverUrl = System.Configuration.ConfigurationManager.AppSettings["ExchangeServerUrl"];
            var putUri = System.Configuration.ConfigurationManager.AppSettings["HttpPutUri"];

            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(o =>
                   {
                       var json = GetCalItems(o, users, serverUrl);
                       if (putUri.Length > 0)
                       {
                           HttpPut(o, json, putUri);
                       }
                       else
                       {
                           Console.WriteLine(json);
                       }
                   });
        }
    }
}
