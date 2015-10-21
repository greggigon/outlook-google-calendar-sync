using System;

namespace OutlookCalendarExporter
{
    class CalendarItem
    {
        public string Subject { get; private set; }
        public DateTime StartTime { get; private set;  }
        public DateTime EndTime { get; private set; } 
        public string Location { get; private set; }
        public string Details { get; private set; }
        public TimeZone TimeZone { get; private set; }
        public bool Busy { get; private set; }
        
        //public List<string> participants = ;

        public static CalendarItem fromDetails(string subject, DateTime startTime, DateTime endTime,
            string location, string details, bool busy)
        {
            return new CalendarItem()
            {
                Subject = subject,
                StartTime = startTime,
                EndTime = endTime,
                TimeZone = TimeZone.CurrentTimeZone,
                Location = location,
                Details = details,
                Busy = busy
            };
        }

    }
}

