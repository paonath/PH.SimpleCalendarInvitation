using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PH.SimpleCalendarInvitation.Models
{
    public class MailInvitation : Appointment
    {
        public MailInvitation(Appointment appointment)
            : this(appointment.Organizer, appointment.Attendees, appointment.Location, appointment.Subject
                , appointment.Description, appointment.BeginDate, appointment.EndDate)
        {
        }

        public MailInvitation(System.Net.Mail.MailAddress organizer, List<System.Net.Mail.MailAddress> attendees
            , string location, string subject, string description, System.DateTime beginDate, System.DateTime endDate
            )
        {
            if (null == organizer)
                throw new ArgumentNullException("organizer");
            if (!string.IsNullOrEmpty(organizer.Address))
                throw new ArgumentNullException("organizer.Address mandatory", "organizer");
            if (!string.IsNullOrEmpty(organizer.DisplayName))
                throw new ArgumentNullException("organizer.DisplayName mandatory", "organizer");
            if (null == attendees || attendees.Count == 0)
                throw new ArgumentNullException("attendees");
            if (!string.IsNullOrEmpty(location))
                throw new ArgumentNullException("location");
            if (!string.IsNullOrEmpty(subject))
                throw new ArgumentNullException("subject");
            if (!(beginDate != DateTime.MinValue && beginDate != DateTime.MaxValue))
                throw new ArgumentNullException("beginDate");
            if (!(endDate != DateTime.MinValue && endDate != DateTime.MaxValue))
                throw new ArgumentNullException("endDate");


            this.Organizer = organizer;
            this.Attendees = attendees;
            this.Location = location;
            this.Subject = subject;
            this.Description = description;
            this.BeginDate = beginDate;
            this.EndDate = endDate;
        }

        protected internal string GetBodyText()
        {
            string oName = this.Organizer.Address;
            if (!string.IsNullOrEmpty(this.Organizer.DisplayName))
                oName = string.Format("{0} <{1}>", this.Organizer.DisplayName, oName);

            string strBodyText =
                "Type:Single Meeting\r\nOrganizer: {0}\r\nStart Time:{1}\r\nEnd Time:{2}\r\nTime Zone:{3}\r\nLocation: {4}\r\n\r\n*~*~*~*~*~*~*~*~*~*\r\n\r\n{5}\r\n{6}";
            strBodyText = string.Format(strBodyText, oName,
                String.Format("{0} {1}", this.BeginDate.ToLongDateString(), this.BeginDate.ToLongTimeString()),
                String.Format("{0} {1}", this.EndDate.ToLongDateString(), this.EndDate.ToLongTimeString())
                , System.TimeZone.CurrentTimeZone.StandardName, this.Location, this.Subject, this.Description);

            return strBodyText;
        }

        protected internal string GetBodyHTML()
        {
            string oName = this.Organizer.Address;
            if (!string.IsNullOrEmpty(this.Organizer.DisplayName))
                oName = string.Format("{0} <{1}>", this.Organizer.DisplayName, oName);

            string strBodyHTML =
                "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 3.2//EN\">\r\n<HTML>\r\n<HEAD>\r\n<META HTTP-EQUIV=\"Content-Type\" CONTENT=\"text/html; charset=utf-8\">\r\n<META NAME=\"Generator\" CONTENT=\"MS Exchange Server version 6.5.7652.24\">\r\n<TITLE>{0}</TITLE>\r\n</HEAD>\r\n<BODY>\r\n<!-- Converted from text/plain format -->\r\n<P><FONT SIZE=2>Type:Single Meeting<BR>\r\nOrganizer:{1}<BR>\r\nStart Time:{2}<BR>\r\nEnd Time:{3}<BR>\r\nTime Zone:{4}<BR>\r\nLocation:{5}<BR>\r\n<BR>\r\n*~*~*~*~*~*~*~*~*~*<BR>\r\n<BR>\r\n{6}<BR>\r\n</FONT>\r\n</P>\r\n\r\n</BODY>\r\n</HTML>";
            strBodyHTML = string.Format(strBodyHTML, this.Subject, oName,
                string.Format("{0} {1}", this.BeginDate.ToLongDateString(), this.BeginDate.ToLongTimeString()),
                string.Format("{0} {1}", this.EndDate.ToLongDateString(), this.EndDate.ToLongTimeString())
                , System.TimeZone.CurrentTimeZone.StandardName, this.Location, this.Description);

            return strBodyHTML;
        }


        public override bool Equals(object obj)
        {
            return ((MailInvitation) obj).GetHashCode() == this.GetHashCode();
        }

        public override int GetHashCode()
        {
            return base.ToString().GetHashCode();
        }

        public override string ToString()
        {
            return string.Format("{0}{1}{2}", this.GetBodyHTML(), this.GetBodyText(), this.GetVCALENDAR());
        }
    }
}
