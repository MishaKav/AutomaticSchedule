using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Newtonsoft.Json;
using System.IO;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Util.Store;

namespace AutomaticSchedule
{
    public static class GoogleApi
    {
        public static List<CalendarListEntry> CalendarsList;
        public static CalendarListEntry SelectedCalendar;
        public static CalendarService CalendarConnection;
        public static GmailService GmailService;
        private const string GooleDateFormat = "yyyy-MM-dd";

        //http://codekicker.de/news/Retrieving-calendar-events-using-Google-Calendar-API
        public static bool ConnectCalendar()
        {
            var secrets = new ClientSecrets
            {
                ClientId = AppSettings.Google.ClientId,
                ClientSecret = AppSettings.Google.ClientSecret
            };

            try
            {
                var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        secrets,
                        new[] { CalendarService.Scope.Calendar },
                        "user",
                        CancellationToken.None)
                .Result;

                var initializer = new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "CheckPoint Work Schedule"
                };
                CalendarConnection = new CalendarService(initializer);
                LoadCalendars();
            }
            catch (Exception ex)
            {
                Utils.WriteErrorLog(ex.Message);
                return false;
            }
            return true;
        }

        #region Calendars Methods
        private static void LoadCalendars()
        {
            if (CalendarConnection.IsNotEmptyObject())
            {
                CalendarsList = CalendarConnection.CalendarList.List().Execute().Items.ToList();

                // take only owners calendars (without readonly)
                CalendarsList.RemoveAll(c => !c.AccessRole.ContainsIgnoreCase("owner"));

                // select primary calendar
                //cbCalendars.SelectedIndex = calendars.FindIndex(c => c.Primary.HasValue && c.Primary.Value);
                SelectedCalendar = CalendarsList.Any(c => c.Summary.ContainsIgnoreCase(AppSettings.Google.DefaultCalendar)) ?
                    CalendarsList.Find(c => c.Summary.ContainsIgnoreCase(AppSettings.Google.DefaultCalendar)) :
                    CalendarsList.FirstOrDefault();
            }
        }

        public static bool AddEvent(string summary, DateTime start, DateTime? end = null, string location = null, string description = null, bool? allDay = null)
        {
            if (SelectedCalendar != null)
            {
                var allDayEvent = allDay.HasValue && allDay.Value;

                var calEvent = new Event
                {
                    Summary = summary,
                    Location = location,
                    Start = new EventDateTime { DateTime = start },
                    End = new EventDateTime { DateTime = end },
                    Description = $"{AppSettings.Google.CalendarIdentifyer}\n\n{description}",
                    Reminders = new Event.RemindersData { UseDefault = false }
                };

                // allDay event or not
                calEvent.Start = allDayEvent ? new EventDateTime { Date = start.ToString(GooleDateFormat) } : new EventDateTime { DateTime = start };

                if (end.HasValue && allDayEvent)
                {
                    calEvent.End = new EventDateTime { Date = end.Value.ToString(GooleDateFormat) };
                }

                //Set Remainder
                if (IsEventExist(calEvent.Summary, calEvent.Start))
                {
                    var prevEvents = GetEvents(calEvent.Summary, calEvent.Start);
                    prevEvents.ForEach(e => DeleteEvent(e.Id));
                }

                CalendarConnection.Events.Insert(calEvent, SelectedCalendar.Id).Execute();
                return IsEventExist(calEvent.Summary, calEvent.Start);
            }
            return false;
        }

        public static bool AddEvent(Reminder reminder)
        {
            if (SelectedCalendar != null)
            {
                var calEvent = new Event
                {
                    Summary = AppSettings.Google.DefaultEventTitle,
                    Location = AppSettings.Google.DefaultLocation,
                    Start = new EventDateTime { DateTime = reminder.Start },
                    End = new EventDateTime { DateTime = reminder.End },
                    Description = $"{AppSettings.Google.CalendarIdentifyer}\nJob: {reminder.JobName}\nHours: {(reminder.End - reminder.Start).TotalHours} h",
                    Reminders = new Event.RemindersData { UseDefault = false }
                };

                //Set Remainder
                if (IsEventExist(calEvent.Summary, calEvent.Start))
                {
                    var prevEvents = GetEvents(calEvent.Summary, calEvent.Start);
                    prevEvents.ForEach(e => DeleteEvent(e.Id));
                }

                CalendarConnection.Events.Insert(calEvent, SelectedCalendar.Id).Execute();
                return IsEventExist(calEvent.Summary, calEvent.Start);
            }
            return false;
        }

        private static List<Event> GetEvents(string eventName, EventDateTime eventDateTime)
        {
            if (SelectedCalendar != null)
            {
                var lr = CalendarConnection.Events.List(SelectedCalendar.Id);

                lr.TimeMin = eventDateTime.DateTime ?? DateTime.ParseExact(eventDateTime.Date, GooleDateFormat, null);
                lr.TimeMax = lr.TimeMin.Value.AddDays(1);

                var request = lr.Execute();

                if (request.IsNotEmptyObject() && request.Items.IsAny())
                {
                    var events = request.Items.ToList();
                    return events.FindAll(e => e.Summary == eventName && e.Description.Contains(AppSettings.Google.CalendarIdentifyer));
                }
            }

            return null;
        }

        private static bool IsEventExist(string eventName, EventDateTime eventDateTime)
        {
            if (SelectedCalendar != null)
            {
                var lr = CalendarConnection.Events.List(SelectedCalendar.Id);

                lr.TimeMin = eventDateTime.DateTime ?? DateTime.ParseExact(eventDateTime.Date, GooleDateFormat, null);
                lr.TimeMax = lr.TimeMin.Value.AddDays(1);

                var request = lr.Execute();

                if (request.IsNotEmptyObject() && request.Items.IsAny())
                {
                    var events = request.Items.ToList();
                    events.RemoveAll(e => !e.Description.ContainsIgnoreCase(AppSettings.Google.CalendarIdentifyer));
                    var isEventExist = events.Any(e => e.Summary.ContainsIgnoreCase(eventName));
                    return isEventExist;
                }
            }

            return false;
        }

        private static void DeleteEvent(string eventId)
        {
            if (SelectedCalendar != null)
            {
                CalendarConnection.Events.Delete(SelectedCalendar.Id, eventId).Execute();
            }
        }

        #endregion Calendars Methods

        #region Direction Services

        public static DirectionsResponse GetDirections(GoogleLocation a, GoogleLocation b)
        {
            var url = $"https://maps.googleapis.com/maps/api/directions/json?origin={a.lat},{a.lng}&destination={b.lat},{b.lng}&mode=driving&units=metric&key={AppSettings.Google.API_KEY}";
            var response = Utils.GetWebRequest(url);

            var res = JsonConvert.DeserializeObject<DirectionsResponse>(response);
            if (res.Status == DirectionsStatus.OK)
                return res;
            else// if (status == DirectionsStatus.ZERO_RESULTS)
                return null; //"Sorry, your search appears to be outside our current coverage area.");
        }

        #endregion Direction Services

        #region Calendar Link Generator
        /// <summary>
        /// Generate link for google calendar
        /// </summary>
        /// <param name="title"></param>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="details"></param>
        /// <returns>link</returns>
        public static string GetGoogleCalendarEvent(string title, DateTime startDate, DateTime? endDate = null, string details = "")
        {
            const string url = "http://www.google.com/calendar/event?action=TEMPLATE";
            const string src = "developer.newconcept@gmail.com";

            if (endDate == null)
            {
                endDate = startDate.AddHours(1);
            }

            var link = $"{url}&" +
                       $"text={title}&" +
                       $"location={AppSettings.Google.DefaultLocation}&" +
                       $"dates={startDate.ToString("yyyyMMddTHHmmss")}/{endDate.Value.ToString("yyyyMMddTHHmmss")}&" +
                       $"details={WebUtility.UrlEncode(AppSettings.Google.CalendarIdentifyer)}%0A{WebUtility.UrlEncode(details)}&" +
                       $"trp=true&sprop=name:Misha Kav&" +
                       $"src={src}";
            return link;
        }
        #endregion

        #region Google API for Gmail
        //https://developers.google.com/gmail/api/v1/reference/users/messages/attachments/get

        public static void CreateCredentinals()
        {
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                var credPath = Path.Combine(Environment.CurrentDirectory, ".credentials/gmail-dotnet-quickstart.json");

                var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    new[] { GmailService.Scope.GmailReadonly },
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

                Debug.WriteLine("Credential file saved to: " + credPath);

                // Create Gmail API service.
                GmailService = new GmailService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Gmail API .NET Automatic Schedule"
                });
            }
        }

        public static void PrintAllLabels()
        {
            // Define parameters of request.
            var request = GmailService.Users.Labels.List("me");

            // List labels.
            var labels = request.Execute().Labels;
            Console.WriteLine("Labels:");
            if (labels != null && labels.Count > 0)
            {
                foreach (var labelItem in labels)
                {
                    Console.WriteLine($"{labelItem.Name}");
                }
            }
            else
            {
                Console.WriteLine("No labels found.");
            }
        }

        /// <summary>
        /// List all Messages of the user's mailbox matching the query.
        /// </summary>
        /// <param name="query">String used to filter Messages returned.</param>
        public static List<Message> ListMessages(string query)
        {
            const string userId = "me";
            var result = new List<Message>();
            var request = GmailService.Users.Messages.List(userId);
            request.Q = query;

            do
            {
                try
                {
                    var response = request.Execute();
                    result.AddRange(response.Messages);
                    request.PageToken = response.NextPageToken;
                }
                catch (Exception e)
                {
                    Utils.WriteErrorLog(e);
                }
            } while (request.PageToken.IsNotNullOrEmpty());

            return result;
        }

        /// <summary>
        /// Get and store attachment from Message with given ID.
        /// </summary>
        /// <param name="messageId">ID of Message containing attachment.</param>
        /// <param name="outputDir">Directory used to store attachments.</param>
        public static List<string> GetAttachments(string messageId, string outputDir)
        {
            const string userId = "me";
            var list = new List<string>();

            try
            {
                var message = GmailService.Users.Messages.Get(userId, messageId).Execute();
                var parts = message.Payload.Parts;
                foreach (var part in parts)
                {
                    if (part.Filename.IsNotNullOrEmpty())
                    {
                        var attId = part.Body.AttachmentId;
                        var attachPart = GmailService.Users.Messages.Attachments.Get(userId, messageId, attId).Execute();

                        // Converting from RFC 4648 base64 to base64url encoding
                        // see http://en.wikipedia.org/wiki/Base64#Implementations_and_history
                        var attachData = attachPart.Data.Replace('-', '+');
                        attachData = attachData.Replace('_', '/');

                        var data = Convert.FromBase64String(attachData);
                        var pathToFile = Path.Combine(outputDir, part.Filename);
                        if (!Directory.Exists(outputDir))
                        {
                            Directory.CreateDirectory(outputDir);
                        }
                        File.WriteAllBytes(pathToFile, data);
                        list.Add(pathToFile);
                    }
                }
            }
            catch (Exception e)
            {
                Utils.WriteErrorLog(e);
            }

            return list;
        }

        #endregion Google API for Gmail
    }


    #region Google autocomplete places class
    public class GooglePlaceResults
    {
        public GoogleStatus status;
        public List<GooglePlaceItem> predictions;
    }

    public class GooglePlaceItem
    {
        public string description;
        public string id;
        public List<GoogleMatchedSubstrings> matched_substrings;
        public string reference;
        public List<GooglePlaceTerms> terms;
        public List<string> types;
    }

    public class GooglePlaceTerms
    {
        public int offset;
        public string value;
    }

    public class GoogleMatchedSubstrings
    {
        public int length;
        public int offset;
    }
    #endregion Google autocomplete places class


    #region Google autocomplete place details class
    public class GooglePlaceDetailsResults
    {
        public List<string> html_attributions;
        public GooglePlaceDetailsItem result;
        public GoogleStatus status;
    }

    public class GooglePlaceDetailsItem
    {
        public List<GoogleAddressComponents> address_components;
        public string adr_address;
        public string formatted_address;
        public GoogleGeometry geometry;
        public string icon;
        public string id;
        public string name;
        public string place_id;
        public string reference;
        public string scope;
        public List<string> types;
        public string url;
        public string vicinity;
    }

    public class GoogleAddressComponents
    {
        public string long_name;
        public string short_name;
        public List<string> types;
    }

    public class GoogleGeometry
    {
        public GoogleLocation location;
    }

    public class GoogleLocation
    {
        public double lat;
        public double lng;
    }
    #endregion Google autocomplete place details class


    #region Google Route Response

    public class DirectionsResponse
    {
        public DirectionsStatus Status { get; set; }

        public List<Route> Routes { get; set; }

        public DirectionsRequest Lc { get; set; }
    }

    public enum DirectionsStatus
    {
        OK,
        NOT_FOUND,
        ZERO_RESULTS,
        MAX_WAYPOINTS_EXCEEDED,
        INVALID_REQUEST,
        OVER_QUERY_LIMIT,
        REQUEST_DENIED,
        UNKNOWN_ERROR,
    }

    public class Route
    {
        public string Summary { get; set; }

        public List<Leg> Legs { get; set; }

        public string Copyrights { get; set; }

        public OverviewPolyline OverviewPolyline { get; set; }

        public List<object> Warnings { get; set; }

        public List<double> WaypointOrder { get; set; }

        public Bounds Bounds { get; set; }
    }

    public class OverviewPolyline
    {
        public string Points { get; set; }
    }

    public class Bounds
    {
        public Location Southwest { get; set; }

        public Location Northeast { get; set; }
    }

    public class Location
    {
        public double Lat { get; set; }

        public double Lng { get; set; }
    }

    public class Leg
    {
        public List<Step> Steps { get; set; }

        public TextValue Duration { get; set; }

        public TextValue Distance { get; set; }

        public Location Start_Location { get; set; }

        public Location End_Location { get; set; }

        public string Start_Address { get; set; }

        public string End_Address { get; set; }
    }

    public class TextValue
    {
        public double Value { get; set; }

        public string Text { get; set; }
    }

    public class Step
    {
        public TravelMode TravelMode { get; set; }

        public Location StartLocation { get; set; }

        public Location EndLocation { get; set; }

        public Polyline polyline { get; set; }

        public TextValue Duration { get; set; }

        public string HtmlInstructions { get; set; }

        public TextValue Distance { get; set; }
    }

    public class Polyline
    {
        public string Points { get; set; }
    }

    public enum TravelMode
    {
        DRIVING,
        BICYCLING,
        TRANSIT,
        WALKING,
    }
    public class DirectionsRequest
    {
        public GoogleLocation Origin { get; set; }

        public GoogleLocation Destination { get; set; }

        public TravelMode TravelMode { get; set; }
    }

    #endregion Google Route Response

    public enum GoogleStatus
    {
        OK,
        NOT_FOUND,
        ZERO_RESULTS,
        MAX_WAYPOINTS_EXCEEDED,
        INVALID_REQUEST,
        OVER_QUERY_LIMIT,
        REQUEST_DENIED,
        UNKNOWN_ERROR
    }
}
