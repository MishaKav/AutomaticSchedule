using System;

namespace AutomaticSchedule
{
    public class TaglitMeeting
    {
        public string GroupCode { get; set; }
        public string GroupCountry { get; set; }
        public int Participants { get; set; }
        public string OrganizerName { get; set; }
        public DateTime MifgashimStart { get; set; }
        public DateTime MifgashimEnd { get; set; }
        public string ConferenceLocation { get; set; }
        public DateTime ConferenceDate { get; set; }
        public string ConferenceTime { get; set; }
        public string GroupSubType { get; set; }
        public string GroupCampus { get; set; }
        public string GroupCommunity { get; set; }
        public string ParameterComments { get; set; }
        public string PilotType { get; set; }
        public string Languages { get; set; }
        public string Cities { get; set; }
        public string Occupations { get; set; }
        public int AgeFrom { get; set; }
        public int AgeTo { get; set; }

        public override string ToString()
        {
            return $"{GroupCode} - {Participants} People, {MifgashimStart:dd/MM/yy} - {MifgashimEnd:dd/MM/yy} [{(MifgashimEnd - MifgashimStart).TotalDays} days], Age: {AgeFrom} - {AgeTo}, {ParameterComments}";
        }
    }
}