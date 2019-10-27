using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using CommandLine;

// Background information:
// https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/get-started-with-ews-client-applications
// https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-get-appointments-and-meetings-by-using-ews-in-exchange

// This requires nuget packages:
// Install-Package Exchange.WebServices.Managed.Api -Version 2.2.1.2
// Install-Package Newtonsoft.Json -Version 12.0.2
// Install-Package CommandLineParser -Version 2.6.0

namespace GetAppt
{
    class Program
    {
        class Appt
        {
            public string subject;
            public DateTime start;
            public DateTime end;
            public string location;
            public string freebusystatus;
        }

        class UserAppts
        {
            public string username;
            public List<Appt> appointments;
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

            // Initialize values for the start and end times, and the number of appointments to retrieve.
            DateTime startDate = DateTime.Now;
            DateTime endDate = startDate.AddDays(5);
            const int NUM_APPTS = 5;

            // Set the start and end time and number of appointments to retrieve.
            CalendarView cView = new CalendarView(startDate, endDate, NUM_APPTS)
            {
                // Limit the properties returned to the appointment's subject, start time, and end time.
                PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End,
                AppointmentSchema.Location, AppointmentSchema.LegacyFreeBusyStatus)
            };

            List<UserAppts> appts = new List<UserAppts>();
            foreach (var u in users)
            {
                var folderIdFromCalendar = new FolderId(WellKnownFolderName.Calendar, u);

                // Initialize the calendar folder object with only the folder ID. 
                CalendarFolder calendar = CalendarFolder.Bind(service, folderIdFromCalendar, new PropertySet());

                // Retrieve a collection of appointments by using the calendar view.
                FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);

                var ua = new UserAppts { username = u, appointments = new List<Appt>() };
                appts.Add(ua);
                foreach (Appointment a in appointments)
                {
                    // Only access the initialized fields in 'a':
                    var obj = new Appt
                    {
                        subject = a.Subject,
                        start = a.Start,
                        end = a.End,
                        location = a.Location,
                        freebusystatus = a.LegacyFreeBusyStatus.ToString()
                    };
                    ua.appointments.Add(obj);
                }
            }
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(appts, Newtonsoft.Json.Formatting.Indented);
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
                    using (var rxStream = res.GetResponseStream()) {
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
