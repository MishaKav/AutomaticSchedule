using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Newtonsoft.Json;

namespace AutomaticSchedule
{
    public static class GoogleApi
    {
        public static List<CalendarListEntry> CalendarsList;
        public static CalendarListEntry SelectedCalendar;
        public static CalendarService CalendarConnection;

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

        public static bool AddEvent(Reminder reminder)
        {
            if (SelectedCalendar != null)
            {
                var calEvent = new Event
                {
                    Summary = "Работа",
                    Location = "Check Point Software Technologies, הסוללים 5, Tel Aviv-Yafo, 67897, Israel",
                    Start = new EventDateTime
                    {
                        DateTime = reminder.Start
                    },
                    End = new EventDateTime
                    {
                        DateTime = reminder.End
                    },
                    Description = $"{AppSettings.Google.CalendarIdentifyer}\nJob: {reminder.JobName}\nHours: {(reminder.End - reminder.Start).TotalHours} h",
                    Reminders = new Event.RemindersData { UseDefault = false }
                };

                //Set Remainder
                if (IsEventExist(calEvent.Summary, calEvent.Start.DateTime))
                {
                    var prevEvents = GetEvents(calEvent.Summary, calEvent.Start.DateTime);
                    prevEvents.ForEach(e => DeleteEvent(e.Id));
                }

                CalendarConnection.Events.Insert(calEvent, SelectedCalendar.Id).Execute();
                return IsEventExist(calEvent.Summary, calEvent.Start.DateTime);
            }
            return false;
        }

        private static List<Event> GetEvents(string eventName, DateTime? datetime)
        {
            if (SelectedCalendar != null)
            {
                var lr = CalendarConnection.Events.List(SelectedCalendar.Id);

                lr.TimeMin = datetime;
                if (datetime != null)
                {
                    lr.TimeMax = datetime.Value.AddDays(1);
                }

                var request = lr.Execute();

                if (request.IsNotEmptyObject() && request.Items.IsAny())
                {
                    var events = request.Items.ToList();
                    return events.FindAll(e => e.Summary == eventName && e.Description.Contains(AppSettings.Google.CalendarIdentifyer));
                }
            }

            return null;
        }

        private static bool IsEventExist(string eventName, DateTime? datetime)
        {
            if (SelectedCalendar != null)
            {
                var lr = CalendarConnection.Events.List(SelectedCalendar.Id);

                lr.TimeMin = datetime;
                if (datetime != null)
                {
                    lr.TimeMax = datetime.Value.AddDays(1);
                }

                var request = lr.Execute();

                if (request.IsNotEmptyObject() && request.Items.IsAny())
                {
                    var events = request.Items.ToList();
                    return events.Any(e => e.Summary.ContainsIgnoreCase(eventName) && e.Description.ContainsIgnoreCase(AppSettings.Google.CalendarIdentifyer));
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
