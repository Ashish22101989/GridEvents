using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using Microsoft.Crm.Sdk.Messages;

namespace AppD365Connection
{

    public class AppointmentEntity
    {

        public Guid Id { get; set; }

        //subject
        public string Subject { get; set; }

        //scheduledstart
        public DateTime StartTime { get; set; }

        //scheduledend
        public DateTime EndTime { get; set; }


        //actualstart
        public DateTime ActualStartTime { get; set; }

        //actualend
        public DateTime ActualEndTime { get; set; }

        //requiredattendees
        public EntityCollection RequiredAttendees { get; set; }

        public EntityCollection OptionalAttendees { get; set; }

        //location
        public string Location { get; set; }

        public EntityReference Owner { get; set; }

        public EntityReference FromParty { get; set; }

        public EntityReference ToParty { get; set; }

        //description
        public string Description { get; set; }

        //regardingobjectid
        public EntityReference Regarding { get; set; }
    }

    public class AppointmentInvite : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.  
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.  
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.  
                Entity entity = (Entity)context.InputParameters["Target"];

                // Obtain the organization service reference which you will need for  
                // web service calls.  
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    if (entity != null && entity.LogicalName == "appointment" && entity.Attributes.Contains("new_sendemail"))
                    {
                        bool sendEmail = (bool)entity.Attributes["new_sendemail"];
                        AppointmentEntity appointment = new AppointmentEntity();

                        Guid emailId = CreateEmail(entity, appointment, service);
                        CreateICSAttachment(emailId, appointment, service);
                        SendEmail(emailId, service);

                    }



                }

                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in FollowUpPlugin.", ex);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("Appointment Invite: {0}", ex.ToString());
                    throw;
                }
            }
        }

        public AppointmentEntity SetAppointment(Entity entity)
        {

            if (entity.Attributes.Contains("scheduledstart") && entity.Attributes["scheduledstart"] != null)
            {

            }

            if (entity.Attributes.Contains("scheduledend") && entity.Attributes["scheduledend"] != null)
            {

            }
            if (entity.Attributes.Contains("description") && entity.Attributes["description"] != null)
            {

            }
            if (entity.Attributes.Contains("") && entity.Attributes[""] != null)
            {

            }
            if (entity.Attributes.Contains("") && entity.Attributes[""] != null)
            {

            }
            if (entity.Attributes.Contains("") && entity.Attributes[""] != null)
            {

            }
            if (entity.Attributes.Contains("") && entity.Attributes[""] != null)
            {

            }
            if (entity.Attributes.Contains("") && entity.Attributes[""] != null)
            {

            }
            if (entity.Attributes.Contains("") && entity.Attributes[""] != null)
            {

            }
            if (entity.Attributes.Contains("") && entity.Attributes[""] != null)
            {

            }
            if (entity.Attributes.Contains("") && entity.Attributes[""] != null)
            {

            }
            if (entity.Attributes.Contains("") && entity.Attributes[""] != null)
            {

            }
            if (entity.Attributes.Contains("") && entity.Attributes[""] != null)
            {


            }
            if (entity.Attributes.Contains("") && entity.Attributes[""] != null)
            {

            }

            return null;
        }

        public Guid CreateEmail(Entity entity, AppointmentEntity appointment, IOrganizationService organizationService)
        {
            try
            {
                Entity email = new Entity("email");
                if (entity.Attributes.Contains("subject") && entity.Attributes["subject"] != null)
                {
                    email["subject"] = entity.Attributes["subject"].ToString();
                    appointment.Subject = entity.Attributes["subject"].ToString();
                }
                if (entity.Attributes.Contains("scheduledstart") && entity.Attributes["scheduledstart"] != null)
                {
                    appointment.StartTime = (DateTime)entity.Attributes["scheduledstart"];
                }
                if (entity.Attributes.Contains("scheduledend") && entity.Attributes["scheduledend"] != null)
                {
                    appointment.EndTime = (DateTime)entity.Attributes["scheduledend"];
                }
                if (entity.Attributes.Contains("description") && entity.Attributes["description"] != null)
                {
                    email["description"] = entity.Attributes["description"].ToString();
                    appointment.Description = entity.Attributes["description"].ToString();
                }
                if (entity.Attributes.Contains("regardingobjectid") && entity.Attributes["regardingobjectid"] != null)
                {
                    email["regardingobjectid"] = (EntityReference)entity.Attributes["regardingobjectid"];
                    appointment.Regarding = (EntityReference)entity.Attributes["regardingobjectid"];
                }
                if (entity.Attributes.Contains("location") && entity.Attributes["location"] != null)
                {
                    appointment.Location = entity.Attributes["location"].ToString();
                }
                if (entity.Attributes.Contains("ownerid") && entity.Attributes["ownerid"] != null)
                {
                    EntityReference owner = (EntityReference)entity.Attributes["ownerid"];
                    Entity fromActivityParty = new Entity("activityparty");
                    fromActivityParty["partyid"] = new EntityReference("systemuser", owner.Id);
                    email["from"] = new Entity[] { fromActivityParty };
                    appointment.Owner = owner;
                }

                EntityCollection requiredAttendees = null;
                if (entity.Attributes.Contains("requiredattendees") && entity.Attributes["requiredattendees"] != null)
                {
                    requiredAttendees = (EntityCollection)entity.Attributes["requiredattendees"];
                    List<Entity> requiredAttendeesList = new List<Entity>();

                    if (requiredAttendees != null && requiredAttendees.Entities.Count > 0)
                    {
                        foreach (var attendee in requiredAttendees.Entities)
                        {
                            if (attendee != null && attendee.Contains("partyid") && attendee.Attributes["partyid"] != null)
                            {
                                Entity toActivityParty = new Entity("activityparty");
                                toActivityParty["partyid"] = attendee.Attributes["partyid"];
                                requiredAttendeesList.Add(toActivityParty);
                            }
                        }
                        email["to"] = requiredAttendeesList.ToArray();
                    }

                    appointment.RequiredAttendees = requiredAttendees;
                }
                if (entity.Attributes.Contains("optionalattendees") && entity.Attributes["optionalattendees"] != null)
                {
                    EntityCollection optionalAttendees = (EntityCollection)entity.Attributes["optionalattendees"];
                    List<Entity> optionalAttendeesList = new List<Entity>();

                    if (optionalAttendees != null && optionalAttendees.Entities.Count > 0)
                    {
                        var optionalAttendeeCollection = (from entity1 in optionalAttendees.Entities
                                                          where !requiredAttendees.Entities.Any(x => ((EntityReference)x.Attributes["partyid"]).Id == (((EntityReference)entity1.Attributes["partyid"]).Id))
                                                          select entity1).ToList();
                        foreach (var attendee in optionalAttendeeCollection)
                        {
                            if (attendee != null && attendee.Contains("partyid") && attendee.Attributes["partyid"] != null)
                            {
                                Entity toActivityParty = new Entity("activityparty");
                                toActivityParty["partyid"] = attendee.Attributes["partyid"];
                                optionalAttendeesList.Add(toActivityParty);
                            }
                        }                        
                            email["cc"] = optionalAttendeesList.ToArray();                       
                    }
                    appointment.OptionalAttendees = optionalAttendees;
                }

                Console.WriteLine("Create Email");
                email["directioncode"] = true;
                return organizationService.Create(email);

            }
            catch (Exception ex)         
            {
                throw new InvalidPluginExecutionException(ex.ToString());
            }
        }



        private void GetAttendees(AppointmentEntity appointment, IOrganizationService organizationService)
        {

        }

        public void CreateICSAttachment(Guid emailId, AppointmentEntity appointment, IOrganizationService organizationService)
        {
            //addressused

            //if (addressused contains data)
            //{
            //    then using partyid attribute get the Name of the account, contact or user
            //}

            StringBuilder requiredAttendeesStr = new StringBuilder();
            StringBuilder optionalAttendeesStr = new StringBuilder();
            if (appointment.RequiredAttendees!= null && appointment.RequiredAttendees.Entities.Count>0)
            {
                EntityCollection requiredAttendeesCollection = appointment.RequiredAttendees;
                RequiredAttendees requiredAttendees = new RequiredAttendees();
                OptionalAttendees optionalAttendees = new OptionalAttendees();                                
                requiredAttendees.AddToList(requiredAttendeesCollection);   
                
                if (appointment.OptionalAttendees!= null && appointment.OptionalAttendees.Entities.Count>0)
                {
                    var optionalAttendeeCollection = (from entity1 in appointment.OptionalAttendees.Entities
                                                      where !appointment.RequiredAttendees.Entities.Any(x => ((EntityReference)x.Attributes["partyid"]).Id == (((EntityReference)entity1.Attributes["partyid"]).Id))
                                                      select entity1).ToList();

                    EntityCollection optionalAttendeesCollection = new EntityCollection();
                    foreach (var entity in optionalAttendeeCollection)
                    {
                        optionalAttendeesCollection.Entities.Add(entity);
                    }
                    optionalAttendees.AddToList(optionalAttendeesCollection);
                }
                
                FetchXmlParameters fetchXmlParameters = new FetchXmlParameters();
                fetchXmlParameters.EntityName = "contact";
                fetchXmlParameters.Attributes = new string[] { "fullname", "contactid", "emailaddress1" };                
                fetchXmlParameters.OrderByAttribute = "fullname";
                fetchXmlParameters.DescendingSortingOrder = false;
                fetchXmlParameters.FilterType = "and";
                fetchXmlParameters.ConditionAttribute = "contactid";
                fetchXmlParameters.ConditionOperator = "in";
                foreach (var contact in requiredAttendees.Contacts)
                {
                    fetchXmlParameters.Conditions.Add(contact.Id.ToString(), contact.Name);
                }

                if (optionalAttendees.Contacts!= null && optionalAttendees.Contacts.Count>0)
                {
                    foreach (var contact in optionalAttendees.Contacts)
                    {
                        fetchXmlParameters.Conditions.Add(contact.Id.ToString(), contact.Name);
                    }
                }

                string contactFetch = CommonUtility.GetFetchXml(fetchXmlParameters);
                EntityCollection entityCollection = CommonUtility.RetrieveMultiple(contactFetch,organizationService);
                

                if (entityCollection != null && entityCollection.Entities.Count > 0)
                {

                    var optionalContactsCollection = (from entity in entityCollection.Entities
                                                      where !requiredAttendees.Contacts.Any(x => x.Id == entity.Id)
                                                      select entity).ToList();

                    var requiredContactsCollection = (from entity in entityCollection.Entities
                                                      where requiredAttendees.Contacts.Any(x => x.Id == entity.Id)
                                                      select entity).ToList();
                    
                    
                    foreach (var entity in requiredContactsCollection)
                    {
                        if (entity.Attributes.Contains("emailaddress1"))
                        {
                            string emailAddress = (string)entity.Attributes["emailaddress1"];
                            requiredAttendeesStr.AppendLine(string.Format("ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;CN=\"{0}\";RSVP=TRUE:mailto:{1}", (string)entity.Attributes["fullname"], emailAddress));
                        }
                    }

                    foreach (var entity in optionalContactsCollection)
                    {
                        if (entity.Attributes.Contains("emailaddress1"))
                        {
                            string emailAddress = (string)entity.Attributes["emailaddress1"];
                            optionalAttendeesStr.AppendLine(string.Format("ATTENDEE;ROLE=OPT-PARTICIPANT;PARTSTAT=NEEDS-ACTION;CN=\"{0}\";RSVP=TRUE:mailto:{1}", (string)entity.Attributes["fullname"], emailAddress));
                        }
                    }

                }
            }

            var ical = new Entity("activitymimeattachment");
            ical["mimetype"] = "text/calendar; charset=UTF-8; method=REQUEST";//;method=REQUEST
            ical["body"] = GetICS(appointment, requiredAttendeesStr, optionalAttendeesStr);//Base64 encoded version of vCalendar above
            ical["subject"] = "Appointment";
            ical["filename"] = "Meeting";
            ical["objectid"] = new EntityReference("email", emailId);
            ical["objecttypecode"] = "email";
            Console.WriteLine("Create Attachment");
            organizationService.Create(ical);
        }

        public string GetICS(AppointmentEntity appointment,StringBuilder requiredAttendees, StringBuilder optionalAttendees)
        {
            StringBuilder str = new StringBuilder();
            str.AppendLine("BEGIN:VCALENDAR");
            str.AppendLine("PRODID:-//Schedule a Meeting");
            str.AppendLine("VERSION:2.0");
            str.AppendLine("METHOD:REQUEST");
            str.AppendLine("BEGIN:VEVENT");
            str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", appointment.StartTime.ToUniversalTime()));//DateTime.Now.AddMinutes(+10)
            str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", DateTime.UtcNow));
            str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", appointment.EndTime.ToUniversalTime())); //DateTime.Now.AddMinutes(+60)
            str.AppendLine("LOCATION: " + appointment.Location);
            str.AppendLine(string.Format("UID:{0}", Guid.NewGuid()));
            // str.AppendLine(string.Format("DESCRIPTION:{0}", "Test meeting"));
            str.AppendLine(string.Format("X-ALT-DESC;FMTTYPE=text/html:{0}", appointment.Description));//msg.Body
            str.AppendLine(string.Format("SUMMARY:{0}", appointment.Subject));//msg.Subject
            str.AppendLine(requiredAttendees.ToString());
            str.AppendLine(optionalAttendees.ToString());
            //str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", appointment.RequiredAttendees));//msg.From.Address"ashish22101989@gmail.com"
            //str.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE:mailto:{1}", "Rohit", "rohit@gmail.com"));
            str.AppendLine("BEGIN:VALARM");
            str.AppendLine("TRIGGER:-PT15M");
            str.AppendLine("ACTION:DISPLAY");
            str.AppendLine("DESCRIPTION:Reminder");
            str.AppendLine("END:VALARM");
            str.AppendLine("END:VEVENT");
            str.AppendLine("END:VCALENDAR");
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(str.ToString());
            return System.Convert.ToBase64String(plainTextBytes);
        }

        public string GetEmailBody(IOrganizationService organizationService)
        {
            return string.Empty;
        }

        public void SendEmail(Guid emailId, IOrganizationService organizationService)
        {
            SendEmailRequest sendEmailreq = new SendEmailRequest
            {
                EmailId = emailId,
                TrackingToken = "",
                IssueSend = true
            };

            //Console.WriteLine("Send Email");
            //SendEmailResponse sendEmailresp = (SendEmailResponse)organizationService.Execute(sendEmailreq);

        }

    }
}
