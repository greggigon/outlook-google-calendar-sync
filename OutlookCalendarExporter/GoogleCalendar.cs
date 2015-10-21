using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Threading;

namespace OutlookCalendarExporter
{
    class GoogleCalendar
    {

        private string apiCredentialsFileName;

        public GoogleCalendar(string apiCredentialsFileName)
        {
            this.apiCredentialsFileName = apiCredentialsFileName;
        }

        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "Google Calendar Outlook Sync";

        UserCredential credential;

        private CalendarService _GetCalendarService()
        {
            string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            credPath = Path.Combine(credPath, ".credentials/google-calendar-sync");

            using (var stream =
                new FileStream(apiCredentialsFileName, FileMode.Open, FileAccess.Read))
            {

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            return new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        public Dictionary<string,CalendarListEntry>  GetListOfCalendars()
        {
            var service = _GetCalendarService();
            var calendarList = service.CalendarList.List().Execute();
            return calendarList.Items.ToDictionary(calendar => calendar.Id);
        }

        public string GetCalendarIdByName(string name)
        {
            var service = _GetCalendarService();

            var result = service.CalendarList.List().Execute();
            var found = result.Items.Where(i => i.Summary == name).ToList();
            
            if (found.Any())
            {
                return found.First().Id;
            }
            return String.Empty;
        }

        public IEnumerable<Event> GetItemsForCalendar(string calendarId, DateTime timeFrom, DateTime timeTo)
        {
            var service = _GetCalendarService();
            var listRequest = service.Events.List(calendarId);

            listRequest.TimeMin = timeFrom;
            listRequest.TimeMax = timeTo;

            return listRequest.Execute().Items;
        }

        public Dictionary<string, List<Event>> CreateEvents(string calendarId, IEnumerable<CalendarItem> eventsToCreate)
        {
            var service = _GetCalendarService();
            var eventsToBeCreated = eventsToCreate.Select(eventFromCalendarItem);
            var created = new List<Event>();
            var errored = new List<Event>();

            foreach (var eventToBeCreated in eventsToBeCreated)
            {
                try {
                    Console.WriteLine(string.Format("    -- Creating EVENT [{0}] for dates [{1}]"
                        ,eventToBeCreated.Summary, eventToBeCreated.Start.DateTime.Value));
                    var result = service.Events.Insert(eventToBeCreated, calendarId).Execute();
                    if (result.Id != null && result.Id != "")
                    {
                        created.Add(result);
                    }
                }catch(Exception ex)
                {
                    Console.WriteLine(ex);
                    errored.Add(eventToBeCreated);
                }
            }
            return new Dictionary<string, List<Event>>()
            {
                { "created", created },
                { "errored", errored}
            };
        }

        public IEnumerable<Event> DeleteEvents(string calendarId, IEnumerable<Event> toBeRemoved)
        {
            var deleted = new List<Event>();
            var service = _GetCalendarService();
            foreach(var toBeDeletedEvent in toBeRemoved)
            {
                Console.WriteLine(string.Format("--> Removing [{0} - {1}]", toBeDeletedEvent.Id, toBeDeletedEvent.Summary));

                var result = service.Events.Delete(calendarId, toBeDeletedEvent.Id).Execute();
                if (string.IsNullOrEmpty(result))
                {
                    deleted.Add(toBeDeletedEvent);
                }
            }
            return deleted;
        }

        public Tuple<IEnumerable<Event>, IEnumerable<Tuple<string, CalendarItem>>> UpdateEvents(string calendarId, IEnumerable<Tuple<string, CalendarItem>> toUpdate)
        {
            var service = _GetCalendarService();
            var updated = new List<Event>();
            var errored = new List<Tuple<string, CalendarItem>>();

            foreach(var eventToUpdate in toUpdate)
            {
                var eventId = eventToUpdate.Item1;
                var eventBody = eventFromCalendarItem(eventToUpdate.Item2);

                Console.WriteLine(" --> Updaing event [{0}]", eventToUpdate.Item2.Subject);
                try {
                    var result = service.Events.Update(eventBody, calendarId, eventId).Execute();
                    updated.Add(result);
                }catch(Exception ex)
                {
                    Console.WriteLine(ex);
                    errored.Add(eventToUpdate);
                }
            }
            return new Tuple<IEnumerable<Event>, IEnumerable<Tuple<string, CalendarItem>>>(updated, errored);
        }

        public Event eventFromCalendarItem(CalendarItem item)
        {
            var e = new Event();
            var startTime = new EventDateTime();
            startTime.DateTime = item.StartTime;
            var endTime = new EventDateTime();
            endTime.DateTime = item.EndTime;

            e.Summary = item.Subject;
            e.Start = startTime;
            e.End = endTime;
            e.Location = item.Location;
            e.Transparency = item.Busy ? "opaque" : "transparent";

            return e;
        }
    }
}
