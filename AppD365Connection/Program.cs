using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Net.Http.Headers;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;

namespace AppD365Connection
{

    class Program
    {
        static void Main(string[] args)
        {

            string connectionStr = @"AuthType=ClientSecret;url=https://orgf86c4a8e.crm8.dynamics.com;ClientId=fa633876-9a9d-48e9-aba8-0e30e828bbb7;ClientSecret=O9CghxySD2~-X~JbPhJZ0220oYLfT5Kr~d";
            CrmServiceClient svc = new CrmServiceClient(connectionStr);
            if (svc.IsReady)
            {
                //appointment&id=f6682d3e-68e1-eb11-bacb-0022486eae28
                //var appointmentEnt = new Entity("appointment");
                //appointmentEnt.Id = new Guid("f6682d3e-68e1-eb11-bacb-0022486eae28");
                //var appointmentEnt = svc.Retrieve("appointment", new Guid("f6682d3e-68e1-eb11-bacb-0022486eae28"), new ColumnSet(true));                

                //var email = new Entity("email")
                //{
                //    To = new ActivityParty[] { toParty, },
                //    From = new ActivityParty[] { fromParty },
                //    Subject = "Test Appointment",
                //    Description = "Please open the attached file to accept the appointment.",
                //    DirectionCode = true,
                //    MimeType = "multipart/mixed"
                //};

                Entity fromActivityParty = new Entity("activityparty");
                Entity toActivityParty = new Entity("activityparty");

                //contact&id=
                Guid contactId = new Guid("25a17064-1ae7-e611-80f4-e0071b661f01");

                fromActivityParty["partyid"] = new EntityReference("systemuser",new Guid("3f715aa5-e2e0-eb11-bacb-0022486ea6e3"));
                toActivityParty["partyid"] = new EntityReference("contact", contactId);
                Console.WriteLine("Create Email");
                Entity email = new Entity("email");
                email["from"] = new Entity[] { fromActivityParty };
                email["to"] = new Entity[] { toActivityParty };
                email["regardingobjectid"] = new EntityReference("contact", contactId);
                email["subject"] = "This is the subject";
                email["description"] = "This is the description.";
                email["directioncode"] = true;
                Guid emailId = svc.Create(email);
          
                //Guid gEmail = svc.Create(email);

                var ical = new Entity("activitymimeattachment");

                ical["mimetype"] = "text/calendar; charset=UTF-8; method=REQUEST";//;method=REQUEST
                ical["body"] = GetICS();//Base64 encoded version of vCalendar above
                ical["subject"] = "Appointment";
                ical["filename"] = "Meeting1";
                ical["objectid"] = new EntityReference("email", emailId);
                ical["objecttypecode"] = email.LogicalName;

                Console.WriteLine("Create Attachment");
                svc.Create(ical);

                SendEmailRequest sendEmailreq = new SendEmailRequest
                {
                    EmailId = emailId,
                    TrackingToken = "",
                    IssueSend = true
                };

                Console.WriteLine("Send Email");
                //SendEmailResponse sendEmailresp = (SendEmailResponse)svc.Execute(sendEmailreq);

            }
        }


        public static string GetICS()
        {

            StringBuilder str = new StringBuilder();
            str.AppendLine("BEGIN:VCALENDAR");
            str.AppendLine("PRODID:-//Schedule a Meeting");
            str.AppendLine("VERSION:2.0");
            str.AppendLine("METHOD:REQUEST");
            str.AppendLine("BEGIN:VEVENT");
            str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", DateTime.Now.AddMinutes(+10)));
            str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", DateTime.UtcNow));
            str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", DateTime.Now.AddMinutes(+60)));
            str.AppendLine("LOCATION: CGI");
            str.AppendLine(string.Format("UID:{0}", Guid.NewGuid()));
            str.AppendLine(string.Format("DESCRIPTION:{0}", "Test meeting"));
            str.AppendLine(string.Format("X-ALT-DESC;FMTTYPE=text/html:{0}","mail body test" ));//msg.Body
            str.AppendLine(string.Format("SUMMARY:{0}", "Metting test title" ));//msg.Subject
            str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", "ashish22101989@gmail.com"));//msg.From.Address

            str.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE:mailto:{1}", "Rohit", "rohit@gmail.com"));

            str.AppendLine("BEGIN:VALARM");
            str.AppendLine("TRIGGER:-PT15M");
            str.AppendLine("ACTION:DISPLAY");
            str.AppendLine("DESCRIPTION:Reminder");
            str.AppendLine("END:VALARM");
            str.AppendLine("END:VEVENT");
            str.AppendLine("END:VCALENDAR");
            //string plainText = "BEGIN:VCALENDAR" +
            //        "PRODID:-//Our Software//CalMan//EN" +
            //        "VERSION:2.0" +
            //        "CALSCALE:GREGORIAN" +
            //        "METHOD:REQUEST" +
            //        "X-MS-OLK-FORCEINSPECTOROPEN:TRUE" +
            //        "BEGIN:VEVENT" +
            //        "DTSTART:20190406T133000Z" +
            //        "DTEND:20190406T143000Z" +
            //        "DTSTAMP:20190402T130711Z" +
            //        "ORGANIZER;CN=The Organizer:mailto:me@us.com" +
            //        "UID:bfnuadapfgs91eqihepacgq2lm@us.com" +
            //        "ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=ACCEPTED;RSVP=TRUE;CN=The Organizer;X-NUM-GUESTS=0:mailto:me@us.com" +
            //        "ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=the recipient;X-NUM-GUESTS=0:mailto:you@them.com" +
            //        "CREATED:20190402T130709Z" +
            //        "DESCRIPTION:some description" +
            //        "LAST-MODIFIED:20190402T130709Z" +
            //        "LOCATION:somehwere" +
            //        "X-ALT-DESC;FMTTYPE=text/html:<html><body><a href='http://bing.com'>Bing</a></body></html>" +
            //        "SEQUENCE:0" +
            //        "STATUS:CONFIRMED" +
            //        "SUMMARY:test" +
            //        "TRANSP:OPAQUE" +
            //        "END:VEVENT" +
            //        "END:VCALENDAR";



            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(str.ToString());
            return System.Convert.ToBase64String(plainTextBytes);            
        }

    }


    

}
 