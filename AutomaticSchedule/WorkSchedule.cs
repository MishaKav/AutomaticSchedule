using System;
using System.Collections.Generic;

namespace AutomaticSchedule
{
    public class WorkSchedule
    {
        public WorkSchedule()
        {
            Name = "Default";
            Reminders = new List<Reminder>();
        }
        public string Name { get; set; }
        public DateTime StartDateTime { get; set; }
        public DateTime EndDateTime { get; set; }
        public List<Reminder> Reminders { get; set; }
    }

    public class Reminder
    {
        public string DayDesc { get; set; }
        public string JobName { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }

        public override string ToString()
        {
            return $"{Start.ToString("ddd")} {JobName}: {Start.ToString("dd/MM/yy HH:mm")} - {End.ToString("HH:mm")}";
        }
    }

    public static class ExtensionWorkSchedule
    {
        public static string ToDisplayFormat(this Reminder source)
        {
            if (source == null) return string.Empty;

            var text = $"{source.Start.ToDefaultDateFormat()}  {source.Start.ToString("ddd"),-8}: {source.Start.ToDefaultTimeFormat()} - {source.End.ToDefaultTimeFormat()}    [{(source.End - source.Start).TotalHours,-2} h]";
            return text;
        }
    }
}
