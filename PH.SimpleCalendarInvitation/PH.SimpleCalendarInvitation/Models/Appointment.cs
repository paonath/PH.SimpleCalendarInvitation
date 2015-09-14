using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace PH.SimpleCalendarInvitation.Models
{
    public class Appointment
    {

        public System.Net.Mail.MailAddress Organizer { get; set; }
        public List<System.Net.Mail.MailAddress> Attendees { get; set; }

        public string Location { get; set; }
        public string Subject { get; set; }
        public string Description { get; set; }
        public System.DateTime BeginDate { get; set; }
        public System.DateTime EndDate { get; set; }


        public string VCALENDAR { get { return this.GetVCALENDAR(); } }


        public const string CALDATEFORMAT = "yyyyMMddTHHmmssZ";

        private const string _UID = "9E37F25B-6A0F-4CE8-847C-B96BE3C8B8A4";


        public string FullUID { get { return GetFullUid(); } }

        public string PartialID { get { return string.Format("{0} {1}", _UID, Assembly.GetExecutingAssembly().FullName); } }

        private string GetFullUid()
        {
            try
            {
                string xhash = this.Organizer.GetHashCode().ToString();
                if (this.Attendees.Count > 0)
                    xhash += this.Attendees.GetHashCode().ToString();

                string x = string.Format("{0} {1}", this.Subject, this.Description);
                xhash += x.GetHashCode();

                return string.Format("{0} {1} {2} {3}", this.PartialID
                    , this.BeginDate.ToUniversalTime().ToString(CALDATEFORMAT),
                    this.EndDate.ToUniversalTime().ToString(CALDATEFORMAT), xhash);
            }
            catch
            {
                return string.Format("{0} {1} {2} {3}", this.PartialID
                        , this.BeginDate.ToUniversalTime().ToString(CALDATEFORMAT),
                        this.EndDate.ToUniversalTime().ToString(CALDATEFORMAT), Guid.NewGuid());
            }

        }





        public override bool Equals(object obj)
        {
            return ((Appointment)obj).GetHashCode() == this.GetHashCode();
        }

        public override int GetHashCode()
        {
            return this.ToString().GetHashCode();
        }

        protected internal string Get_VCALENDAR_For_Organizer()
        {
            return this.VCALENDAR.Replace("METHOD:REQUEST", "METHOD:PUBLISH");
        }

        protected internal string GetVCALENDAR()
        {


            #region old
            //            string organizerCN = string.Format("\"{0}\"" , this.Organizer.DisplayName);

            //            
            ////            string strBodyCalendar = @"BEGIN:VCALENDAR
            ////                                        METHOD:REQUEST
            ////                                        PRODID:Microsoft CDO for Microsoft Exchange
            ////                                        VERSION:2.0
            ////                                        BEGIN:VTIMEZONE
            ////                                        TZID:(GMT-06.00) Central Time (US &amp; Canada)
            ////                                        X-MICROSOFT-CDO-TZID:11
            ////                                        BEGIN:STANDARD
            ////                                        DTSTART:16010101T020000
            ////                                        TZOFFSETFROM:-0500
            ////                                        TZOFFSETTO:-0600
            ////                                        RRULE:FREQ=YEARLY;WKST=MO;INTERVAL=1;BYMONTH=11;BYDAY=1SU
            ////                                        END:STANDARD
            ////                                        BEGIN:DAYLIGHT
            ////                                        DTSTART:16010101T020000
            ////                                        TZOFFSETFROM:-0600
            ////                                        TZOFFSETTO:-0500
            ////                                        RRULE:FREQ=YEARLY;WKST=MO;INTERVAL=1;BYMONTH=3;BYDAY=2SU
            ////                                        END:DAYLIGHT
            ////                                        END:VTIMEZONE
            ////                                        BEGIN:VEVENT
            ////                                        DTSTAMP:{8}
            ////                                        DTSTART:{0}
            ////                                        SUMMARY:{7}
            ////                                        UID:{5}
            ////                                        ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=" + organizerCN + @":MAILTO:{9}
            ////                                        ACTION;RSVP=TRUE;CN=\"{4}\":MAILTO:{4}
            ////                                        ORGANIZER;CN=\"{3}\":mailto:{4}
            ////                                        LOCATION:{2}
            ////                                        DTEND:{1}
            ////                                        DESCRIPTION:{7}\\N
            ////                                        SEQUENCE:1
            ////                                        PRIORITY:5
            ////                                        CLASS:
            ////                                        CREATED:{8}
            ////                                        LAST-MODIFIED:{8}
            ////                                        STATUS:CONFIRMED
            ////                                        TRANSP:OPAQUE
            ////                                        X-MICROSOFT-CDO-BUSYSTATUS:BUSY
            ////                                        X-MICROSOFT-CDO-INSTTYPE:0
            ////                                        X-MICROSOFT-CDO-INTENDEDSTATUS:BUSY
            ////                                        X-MICROSOFT-CDO-ALLDAYEVENT:FALSE
            ////                                        X-MICROSOFT-CDO-IMPORTANCE:1
            ////                                        X-MICROSOFT-CDO-OWNERAPPTID:-1
            ////                                        X-MICROSOFT-CDO-ATTENDEE-CRITICAL-CHANGE:{8}
            ////                                        X-MICROSOFT-CDO-OWNER-CRITICAL-CHANGE:{8}
            ////                                        BEGIN:VALARM
            ////                                        ACTION:DISPLAY
            ////                                        DESCRIPTION:REMINDER
            ////                                        TRIGGER;RELATED=START:-PT00H15M00S
            ////                                        END:VALARM
            ////                                        END:VEVENT
            ////                                        END:VCALENDAR";
            ////            strBodyCalendar = string.Format(strBodyCalendar, this.BeginDate.ToUniversalTime().ToString(strCalDateFormat)
            ////                , this.EndDate.ToUniversalTime().ToString(strCalDateFormat),
            ////            this.Location, this.Organizer.DisplayName, this.Organizer.Address, this.UID.ToString("B"), this.Description, this.Subject,
            ////            DateTime.Now.ToUniversalTime().ToString(strCalDateFormat), this.Attendees.ToString());

            //            string myCalSTR = @"
            //BEGIN:VCALENDAR
            //PRODID: paonath - PH.CalendarAppointment
            //VERSION:2.0
            //METHOD:REQUEST
            //BEGIN:VTIMEZONE
            //TZID:{0}
            //X-TZINFO:{0}
            //BEGIN:DAYLIGHT
            //TZOFFSETTO:+020000
            //TZOFFSETFROM:+010000
            //TZNAME:(DST)
            //DTSTART:20130331T020000
            //RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
            //END:DAYLIGHT
            //BEGIN:STANDARD
            //TZOFFSETTO:+010000
            //TZOFFSETFROM:+020000
            //TZNAME:(STD)
            //DTSTART:20121028T030000
            //RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
            //END:STANDARD
            //END:VTIMEZONE
            //BEGIN:VEVENT
            //ORGANIZER;CN={1}
            //DTSTART:{2}
            //DTEND:{3}
            //{4}
            //UID:{5}
            //SUMMARY:{6}
            //LOCATION:{7}
            //DESCRIPTION:{8}
            //CREATED:{9}
            //LAST-MODIFIED:{10}
            //CLASS:PUBLIC
            //PRIORITY:5
            //SEQUENCE:0
            //STATUS:TENTATIVE
            //TRANSP:OPAQUE
            //DTSTAMP:{11}
            //BEGIN:VALARM
            //TRIGGER:-PT15M
            //ACTION:DISPLAY
            //DESCRIPTION:Reminder
            //END:VALARM
            //END:VEVENT
            //END:VCALENDAR
            //";

            //    string cnOrganizer = string.Format("CN=\"{0}\":mailto:{1}" 
            //        ,this.Organizer.DisplayName , this.Organizer.Address );

            //    StringBuilder sbCnAttendees = new StringBuilder();
            //    sbCnAttendees.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";ROLE=REQ-PARTICIPANT;RSVP=TRUE:mailto:{1}"
            //     , this.Organizer.DisplayName, this.Organizer.Address));
            //    foreach (var item in this.Attendees)
            //    {

            //         sbCnAttendees.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";ROLE=REQ-PARTICIPANT;RSVP=TRUE:mailto:{1}"  
            //             , item.DisplayName , item.Address ));
            //    }
            //    string created = DateTime.Now.ToUniversalTime().ToString(strCalDateFormat);
            //    string attendees = sbCnAttendees.ToString();
            //    string id = this.FullUID;

            //    string tzid = this.TZID;
            //    //tzid = "Europe/Berlin";


            //    string strBodyCalendar = string.Format(myCalSTR
            //        ,tzid , cnOrganizer , this.BeginDate.ToUniversalTime().ToString(strCalDateFormat)
            //        , this.EndDate.ToUniversalTime().ToString(strCalDateFormat), attendees
            //        , id, this.Subject, this.Location, this.Description, created, created, created);


            //            return strBodyCalendar;
            #endregion

            string id = this.PartialID.Replace(" ", "//");
            //METHOD:REQUEST
            //METHOD:PUBLISH

            string googleICS =
@"BEGIN:VCALENDAR
PRODID:-//paonath//" + id + @"
VERSION:2.0
CALSCALE:GREGORIAN
METHOD:REQUEST
BEGIN:VEVENT
DTSTART:" + this.BeginDate.ToUniversalTime().ToString(CALDATEFORMAT) + @"
DTEND:" + this.EndDate.ToUniversalTime().ToString(CALDATEFORMAT) + @"
DTSTAMP:" + DateTime.Now.ToUniversalTime().ToString(CALDATEFORMAT) + @"
ORGANIZER;CN=" + this.Organizer.DisplayName + @":mailto:" + this.Organizer.Address + @"
UID:" + this.FullUID + @"
ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=ACCEPTED;RSVP=TRUE
    ;CN=" + this.Organizer.DisplayName + @";X-NUM-GUESTS=0:mailto:" + this.Organizer.Address;
            foreach (var item in this.Attendees)
            {
                googleICS += @"ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=
    TRUE;CN=" + item.DisplayName + @";X-NUM-GUESTS=0:mailto:" + item.Address + @"";
            }

            googleICS += Environment.NewLine + @"CREATED:" + DateTime.Now.ToUniversalTime().ToString(CALDATEFORMAT) + @"
DESCRIPTION:" + this.Description + @"
LAST-MODIFIED:" + DateTime.Now.ToUniversalTime().ToString(CALDATEFORMAT) + @"
LOCATION:" + this.Location + @"
SEQUENCE:0
STATUS:CONFIRMED
SUMMARY:" + this.Subject + @"
TRANSP:OPAQUE
END:VEVENT
END:VCALENDAR";
            return googleICS;

        }

        public override string ToString()
        {
            return this.VCALENDAR;
        }
    }
}

