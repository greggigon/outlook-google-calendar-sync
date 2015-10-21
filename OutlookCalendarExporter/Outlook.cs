using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Outlook;


namespace OutlookCalendarExporter
{
    class Outlook
    {


        public IEnumerable<CalendarItem> ListOfCalendarItemsFromRange(int daysFromYesterday, bool includeRecuring = false)
        {
            Application OutlookApplication = new Application();
            NameSpace NameSpace = OutlookApplication.GetNamespace("MAPI");
            MAPIFolder folder = NameSpace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

            string filter = "[Start] >='" + DateTime.Now.AddDays(-1).ToString("g") +
                "' AND [End] <= '" + DateTime.Now.AddDays(daysFromYesterday).ToString("g") + "'";
            folder.Items.IncludeRecurrences = includeRecuring;

            Items itemsInDateRange = folder.Items.Restrict(filter);
            itemsInDateRange.IncludeRecurrences = includeRecuring;
            itemsInDateRange.Sort("[Start]", Type.Missing);

            var stuff = itemsInDateRange.Cast<AppointmentItem>();

            return stuff.Where(item => item.RecurrenceState == OlRecurrenceState.olApptNotRecurring).Select(item => CalendarItem.fromDetails(
                item.Subject, item.Start, item.End, item.Location, item.Body, BusyStatus(item.BusyStatus)));
        }

        private bool BusyStatus(OlBusyStatus status)
        {
            return status == OlBusyStatus.olBusy || status == OlBusyStatus.olTentative;
        }
    }

}
