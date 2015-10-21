using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Calendar.v3.Data;
using Quartz;

namespace OutlookCalendarExporter 
{
    class SyncJob : IJob
    {

        public void Execute(IJobExecutionContext context)
        {
            var start = DateTime.Now;
            string calendarId = (string) context.JobDetail.JobDataMap["calendarId"];
            string calendarApiCredentialsFileName = (string) context.JobDetail.JobDataMap["calendarApiCredentialsFileName"];
            int numberOfDaysToSync = (int) context.JobDetail.JobDataMap["numberOfDaysToSync"];

            var googleCalendar = new GoogleCalendar(calendarApiCredentialsFileName);
            var outlook = new Outlook();

            var outlookEvents = outlook.ListOfCalendarItemsFromRange(numberOfDaysToSync, false);
            var googleEvents = googleCalendar.GetItemsForCalendar(calendarId, DateTime.Now.AddDays(-1), DateTime.Now.AddDays(numberOfDaysToSync));

            Console.WriteLine(
                string.Format("----------------------------------\n Events from Outlook [{0}] \nEvents in Google [{1}]\n{2}\n-------------------------------\n\n",
                outlookEvents.Count(), googleEvents.Count(), DateTime.Now.ToLocalTime()));
            

            var eventsToBeDeleted = FilterEventsToBeDeleted(outlookEvents, googleEvents);
            Console.WriteLine(string.Format(" --> To be DELETED [{0}]", eventsToBeDeleted.Count()));
            var resultOfDeletion = googleCalendar.DeleteEvents(calendarId, eventsToBeDeleted);
            Console.WriteLine(string.Format(" --> Removed [{0}] events", resultOfDeletion.Count()));

            var eventsToBeCreated = FilterEventsToBeCreated(outlookEvents, googleEvents);
            Console.WriteLine(string.Format(" --> To be CREATED [{0}]", eventsToBeCreated.Count()));
            var resultOfCreation = googleCalendar.CreateEvents(calendarId, eventsToBeCreated);
            Console.WriteLine(string.Format(" --> CREATED events SUCCESS [{0}]", resultOfCreation["created"].Count()));
            Console.WriteLine(string.Format(" --> CREATED events ERRORS [{0}]", resultOfCreation["errored"].Count()));

            IEnumerable<Tuple<string, CalendarItem>> eventsToBeUpdated = FilterEventsToBeUpdated(outlookEvents, googleEvents);
            var resultOfUpdate = googleCalendar.UpdateEvents(calendarId, eventsToBeUpdated);
            Console.WriteLine(string.Format(" --> UPDATED events SUCCESS [{0}]", resultOfUpdate.Item1.Count()));
            Console.WriteLine(string.Format(" --> UPDATED events ERRORS [{0}]", resultOfUpdate.Item2.Count()));

            Console.WriteLine("=====================================");
            Console.WriteLine(string.Format("Sync Completed. Took [{0}] seconds",
                (DateTime.Now - start).TotalMilliseconds / 1000));
            Console.WriteLine("=====================================");

        }

        private static IEnumerable<Event> FilterEventsToBeDeleted(IEnumerable<CalendarItem> outlookItems, IEnumerable<Event> googleItems)
        {
            var collectionOfOutlookItemsAsMyIds = outlookItems.Select(i => MyId(i.Subject, i.StartTime, i.EndTime));

            return googleItems.Where(e => {
                var myId = MyId(e.Summary, e.Start.DateTime.Value, e.End.DateTime.Value);
                return !collectionOfOutlookItemsAsMyIds.Any(id => id == myId);
            });
        }

        private static IEnumerable<CalendarItem> FilterEventsToBeCreated(IEnumerable<CalendarItem> outlookItems, IEnumerable<Event> googleEvents)
        {
            var collectionOfGoogleIds = googleEvents.Select(g => MyId(g.Summary, g.Start.DateTime.Value, g.End.DateTime.Value));
            return outlookItems.Where(outlookItem =>
            {
                var myOutlookId = MyId(outlookItem.Subject, outlookItem.StartTime, outlookItem.EndTime);
                return !collectionOfGoogleIds.Any(googleId => googleId == myOutlookId);
            });
        }

        private static IEnumerable<Tuple<string, CalendarItem>> FilterEventsToBeUpdated(IEnumerable<CalendarItem> outlookItems, IEnumerable<Event> googleEvents)
        {
            var googleEventByMyId = googleEvents.ToDictionary(g => MyId(g.Summary, g.Start.DateTime.Value, g.End.DateTime.Value));
            var outlookEventsByMyId = outlookItems.ToDictionary(i => MyId(i.Subject, i.StartTime, i.EndTime));

            var myIdsToBePotentialyUpdated = googleEventByMyId.Keys.Intersect(outlookEventsByMyId.Keys);
            return myIdsToBePotentialyUpdated.Where(id => AreEventsDifferent(outlookEventsByMyId[id], googleEventByMyId[id]))
               .Select(id => new Tuple<string, CalendarItem>(googleEventByMyId[id].Id, outlookEventsByMyId[id]));
        }

        private static string MyId(string summary, DateTime start, DateTime end)
        {
            return string.Format("{0}-{1}-{2}", start.ToString("MMddyyyy_HHmm"), end.ToString("MMddyyyy_HHmm"), summary);
        }

        private static bool AreEventsDifferent(CalendarItem outlookEvent, Event googleEvent)
        {
            return outlookEvent.Busy != (_BusyFromGoogle(googleEvent.Transparency))
                || outlookEvent.Location != googleEvent.Location;
        }

        private static bool _BusyFromGoogle(string transparency)
        {
            return transparency == null || transparency == "opaque";
        }
    }
}
