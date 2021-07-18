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
    public class RequiredAttendees
    {
        public RequiredAttendees()
        {
            this.Contacts = new List<Contact>();
            this.Accounts = new List<Account>();
            this.Users = new List<User>();
        }
        public List<Contact> Contacts { get; set; }
        public List<User> Users { get; set; }
        public List<Account> Accounts { get; set; }

        public void AddToList(Contact contact)
        {
            this.Contacts.Add(contact);
        }

        public void AddToList(User user)
        {
            this.Users.Add(user);
        }

        public void AddToList(Account account)
        {
            this.Accounts.Add(account);
        }

        public List<Contact> GetDistinctContactAttendees()
        {
            return Contacts.GroupBy(item => item.EmailAddress).Select(item => item.First()).ToList();
        }

        public List<Account> GetDistinctAccountAttendees()
        {
            return Accounts.GroupBy(item => item.EmailAddress).Select(item => item.First()).ToList();
        }

        public List<User> GetDistinctUserAttendees()
        {
            return Users.GroupBy(item => item.EmailAddress).Select(item => item.First()).ToList();
        }

        public List<Contact> GetDistinctOptionalContacts(List<Contact> optionalContacts, List<Contact> requiredContacts)
        {
            if (optionalContacts.Count==0 || requiredContacts.Count == 0)
            {
                return null;
            }

            var optionalContactsList = (from optContact in optionalContacts
                       where !requiredContacts.Any(
                                         x => x.Id == optContact.Id)
                       select optContact).ToList();

            return optionalContactsList.ToList();

        }

        public List<User> GetDistinctOptionalUsers(List<User> optionalUsers, List<User> requiredUsers)
        {
            if (optionalUsers.Count == 0 || requiredUsers.Count == 0)
            {
                return null;
            }

            var optionalUsersList = (from optUser in optionalUsers
                                        where !requiredUsers.Any(
                                                          x => x.Id == optUser.Id)
                                        select optUser).ToList();

            return optionalUsersList.ToList();

        }

        public List<Account> GetDistinctOptionalAccounts(List<Account> optionalAccounts, List<Account> requiredAccounts)
        {
            if (optionalAccounts.Count == 0 || requiredAccounts.Count == 0)
            {
                return null;
            }

            var optionalAccountsList = (from optAccount in optionalAccounts
                                        where !requiredAccounts.Any(
                                                          x => x.Id == optAccount.Id)
                                        select optAccount).ToList();

            return optionalAccountsList.ToList();

        }

        public void AddToList(EntityCollection attendeeCollection)
        {
            foreach (var attendee in attendeeCollection.Entities)
            {

                if (attendee != null && attendee.Contains("partyid") && attendee.Attributes["partyid"] != null)
                {
                    var entityReference = (EntityReference)attendee.Attributes["partyid"];
                    //ent.Id = entRef.Id
                    if (entityReference.LogicalName == "contact")
                    {
                        int contactCount = this.Contacts.Where(item => item.Id == entityReference.Id).Count();

                        if (contactCount == 0)
                        {
                            this.AddToList(new Contact { Id = entityReference.Id, Name = entityReference.Name });
                        }

                       
                        
                    }
                    else if (entityReference.LogicalName == "account")
                    {

                        int usersCount = this.Users.Where(item => item.Id == entityReference.Id).Count();
                        if (usersCount == 0)
                        {
                            this.AddToList(new User { Id = entityReference.Id, Name = entityReference.Name });
                        }
                    }
                    else if (entityReference.LogicalName == "systemuser")
                    {
                        int accountCount = this.Accounts.Where(item => item.Id == entityReference.Id).Count();
                        if (accountCount == 0)
                        {
                            this.AddToList(new Account { Id = entityReference.Id, Name = entityReference.Name });
                        }
                    }
                }
            }
        }

        public void AddToList(EntityCollection relationshipAttendeeEntityCollection, RequiredAttendees requiredAttendeesCopy)
        {
            foreach (var attendee in relationshipAttendeeEntityCollection.Entities)
            {
                if (attendee != null && attendee.Attributes.Contains("partyid") && attendee.Attributes["partyid"] != null)
                {
                    EntityReference entityReference = (EntityReference)attendee.Attributes["new_contact"];
                    this.AddToList(new Contact { Id = entityReference.Id, Name = entityReference.Name });
                    requiredAttendeesCopy.AddToList(new Contact { Id = entityReference.Id, Name = entityReference.Name });
                }
                else if (attendee != null && attendee.Attributes.Contains("new_user") && attendee.Attributes["new_user"] != null)
                {
                    EntityReference entityReference = (EntityReference)attendee.Attributes["new_user"];

                    this.AddToList(new User { Id = entityReference.Id, Name = entityReference.Name });
                    requiredAttendeesCopy.AddToList(new User { Id = entityReference.Id, Name = entityReference.Name });
                }
                else if (attendee != null && attendee.Attributes.Contains("new_account") && attendee.Attributes["new_account"] != null)
                {
                    EntityReference entityReference = (EntityReference)attendee.Attributes["new_account"];
                    this.AddToList(new Account { Id = entityReference.Id, Name = entityReference.Name });
                    requiredAttendeesCopy.AddToList(new Account { Id = entityReference.Id, Name = entityReference.Name });
                }
            }
        }
    }

    public class FetchXmlParameters
    {

        public FetchXmlParameters()
        {
            this.Conditions = new Dictionary<string, string>();
        }
        public string EntityName { get; set; }

        public string[] Attributes { get; set; }

        public string OrderByAttribute { get; set; }

        public bool DescendingSortingOrder { get; set; }

        public string FilterType { get; set; }

        public string ConditionAttribute { get; set; }

        public string ConditionOperator { get; set; }

        public Dictionary<string, string> Conditions { get; set; }

    }

    public class FetchXml
    {
        public FetchXml(FetchXmlParameters fetchXmlParameters)
        {
            this.Parameter = fetchXmlParameters;
        }

        private FetchXmlParameters Parameter { get; set; }

        public string GetFetchXml()
        {
            StringBuilder fetchXml = new StringBuilder();
            fetchXml.AppendLine("<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">");
            fetchXml.AppendLine("  <entity name=\""+ this.Parameter.EntityName+"\">");

            foreach (var attribute in this.Parameter.Attributes)
            {
                fetchXml.AppendLine("    <attribute name=\"" + attribute + "\" />");
            }

            fetchXml.AppendLine("    <order attribute=\"" + this.Parameter.OrderByAttribute + "\" descending=\""+ this.Parameter.DescendingSortingOrder.ToString().ToLower()+"\" />");


            fetchXml.AppendLine("    <filter type=\""+ this.Parameter.FilterType + "\">");

            fetchXml.AppendLine("      <condition attribute=\""+ this.Parameter.ConditionAttribute+"\" operator=\""+ this.Parameter.ConditionOperator+"\">");

            foreach (var condition in this.Parameter.Conditions)
            {
                //{25A17064-1AE7-E611-80F4-E0071B661F01}
                fetchXml.AppendLine("        <value uitype=\""+ this.Parameter.EntityName + "\">{"+ condition.Key + "}</value>");
            }

            fetchXml.AppendLine("      </condition>");
            fetchXml.AppendLine("    </filter>");
            fetchXml.AppendLine("  </entity>");
            fetchXml.AppendLine("</fetch>");
            return fetchXml.ToString();
        }

    }

    public  static class CommonUtility
    {
        public static Entity Retrieve(Guid guid, string entityName, string[] columns, IOrganizationService service)
        {
            return service.Retrieve(entityName, guid, new ColumnSet(columns));
        }

        public static EntityCollection RetrieveMultiple(string fetchXml,IOrganizationService service)
        {
            return service.RetrieveMultiple(new FetchExpression(fetchXml));
        }

        public static string GetFetchXml(FetchXmlParameters fetchXmlParameters)
        {
            FetchXml fetchXml = new FetchXml(fetchXmlParameters);
            return fetchXml.GetFetchXml();
        }

    }

    public class OptionalAttendees
    {
        public OptionalAttendees()
        {
            this.Contacts = new List<Contact>();
            this.Accounts = new List<Account>();
            this.Users = new List<User>();
        }
        public List<Contact> Contacts { get; set; }
        public List<User> Users { get; set; }
        public List<Account> Accounts { get; set; }

        public void AddToList(EntityCollection attendeeCollection)
        {
            foreach (var attendee in attendeeCollection.Entities)
            {

                if (attendee != null && attendee.Contains("partyid") && attendee.Attributes["partyid"] != null)
                {
                    var entityReference = (EntityReference)attendee.Attributes["partyid"];
                    //ent.Id = entRef.Id
                    if (entityReference.LogicalName == "contact")
                    {
                        int contactCount = this.Contacts.Where(item => item.Id == entityReference.Id).Count();
                        if (contactCount == 0)
                        {
                            this.AddToList(new Contact { Id = entityReference.Id, Name = entityReference.Name });
                        }
                    }
                    else if (entityReference.LogicalName == "account")
                    {

                        int usersCount = this.Users.Where(item => item.Id == entityReference.Id).Count();
                        if (usersCount == 0)
                        {
                            this.AddToList(new User { Id = entityReference.Id, Name = entityReference.Name });
                        }
                    }
                    else if (entityReference.LogicalName == "systemuser")
                    {
                        int accountCount = this.Accounts.Where(item => item.Id == entityReference.Id).Count();
                        if (accountCount == 0)
                        {
                            this.AddToList(new Account { Id = entityReference.Id, Name = entityReference.Name });
                        }
                    }
                }
            }
        }

        private void AddToList(Contact contact)
        {
            this.Contacts.Add(contact);
        }

        private void AddToList(User user)
        {
            this.Users.Add(user);
        }

        private void AddToList(Account account)
        {
            this.Accounts.Add(account);
        }


    }

    public class Contact
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public string EntityName { get { return "contact"; } }
        public string EmailAddress { get; set; }

    }

    public class Account
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public string EntityName { get { return "account"; } }

        public string EmailAddress { get; set; }

        public Contact PrimaryContact { get; set; }
    }

    public class User
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public string EntityName { get { return "systemuser"; } }
        public string EmailAddress { get; set; }
    }

    public class Appointment
    {

        //optionalattendees
        public void OnCreate(IServiceProvider serviceProvider)
        {
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            // Obtain the organization service reference.
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            if (tracingService == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve the tracing service.");
            }

            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {

                    Entity appointment = (Entity)context.InputParameters["Target"];

                    if (appointment != null)
                    {
                        Guid appointmentId = appointment.Id;
                        RelationshipAttendee relationshipAttendee = new RelationshipAttendee();
                        EntityCollection requiredAttendees = null;
                        EntityCollection optionalAttendees = null;
                        if (appointment.Contains("requiredattendees") && appointment.Attributes["requiredattendees"] != null)
                        {
                            requiredAttendees = (EntityCollection)appointment.Attributes["requiredattendees"];
                            relationshipAttendee.Create(requiredAttendees, service);
                        }


                        if (appointment.Contains("optionalattendees") && appointment.Attributes["optionalattendees"] != null)
                        {
                            optionalAttendees = (EntityCollection)appointment.Attributes["optionalattendees"];
                            relationshipAttendee.Create(optionalAttendees, service);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.ToString());
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="serviceProvider"></param>
        public void OnUpdate(IServiceProvider serviceProvider)
        {
            #region pseudo logic
            // When record is updated
            //- 2 cases
            //  - when fields are added or removed. (it will make "Relationship Attendee" entity in sync with required and optionnal attendee fields)
            //     - get Pre image and post image of appointment entity
            //     - if pre image has more records then elemnt is added else element is removed.
            //     if element is added then perform this logic
            //         - get the list of "Relationship Attendee" records and store it in respective objects.		
            //         -iterate over the list of "required attendess" on appointment entity.
            //          - check based on entity logical name and Guid id in "Relationship Attendee" objects.
            //          - If record is not found then create record in "Relationship Attendee" entity.
            //        - Also remove the record from the "Relationship Attendee" objects(create a copy of new "Relationship Attendee" object)
            //          - If the records are present in Objects after iteration is over - check if this record is available in pre image and is not available in post image then
            //          - delete these records from the "Relationship Attendee" entity.
            //        - get the list of optional attendees
            //          - perform the same logic above.
            //if elemnt is removed then perform this logic
            //    - get the list of objects that are removed - get distict from pre image and post image
            //    - iterate over these reocrd and delete it from RA entity
            #endregion pseudo logic

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            // Obtain the organization service reference.
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            if (tracingService == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve the tracing service.");
            }

            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity appointment = (Entity)context.InputParameters["Target"];
                    if (appointment != null)
                    {
                        Guid appointmentId = appointment.Id;
                        RelationshipAttendee relationshipAttendee = new RelationshipAttendee();
                        EntityCollection preImageRequiredAttendees = null;
                        EntityCollection preImageOptionalAttendees = null;
                        EntityCollection postImageRequiredAttendees = null;
                        EntityCollection postImageOptionalAttendees = null;
                        Entity preAttendeesImage = null;
                        Entity postAttendeesImage = null;
                        int preImageRequiredAttendeesCount = 0;
                        int preImageOptionalAttendeesCount = 0;
                        int postImageRequiredAttendeesCount = 0;
                        int postImageOptionalAttendeesCount = 0;

                        #region get pre and post image required and optional attendees and their count

                        if (context.PreEntityImages.Contains("AttendeesImage") && context.PreEntityImages["AttendeesImage"] is Entity)
                        {
                            preAttendeesImage = (Entity)context.PreEntityImages["AttendeesImage"];
                            preImageRequiredAttendees = (EntityCollection)preAttendeesImage.Attributes["requiredattendees"];
                            preImageOptionalAttendees = (EntityCollection)preAttendeesImage.Attributes["optionalattendees"];

                            if (preImageRequiredAttendees != null && preImageRequiredAttendees.Entities.Count > 0)
                            {
                                preImageRequiredAttendeesCount = preImageRequiredAttendees.Entities.Count;
                            }

                            if (preImageOptionalAttendees != null && preImageOptionalAttendees.Entities.Count > 0)
                            {
                                preImageOptionalAttendeesCount = preImageOptionalAttendees.Entities.Count;
                            }
                        }


                        if (context.PreEntityImages.Contains("AttendeesImage") && context.PreEntityImages["AttendeesImage"] is Entity)
                        {
                            postAttendeesImage = (Entity)context.PostEntityImages["AttendeesImage"];
                            postImageRequiredAttendees = (EntityCollection)postAttendeesImage.Attributes["requiredattendees"];
                            postImageOptionalAttendees = (EntityCollection)postAttendeesImage.Attributes["optionalattendees"];

                            if (postImageRequiredAttendees != null && postImageRequiredAttendees.Entities.Count > 0)
                            {
                                postImageRequiredAttendeesCount = postImageRequiredAttendees.Entities.Count;
                            }
                            if (postImageOptionalAttendees != null && postImageOptionalAttendees.Entities.Count > 0)
                            {
                                postImageOptionalAttendeesCount = postImageOptionalAttendees.Entities.Count;
                            }
                        }

                        #endregion get pre and post image required and optional attendees and their count

                        //condition checks that records are added to the required attendee field on apponitment entity and therfore we need to create records in RA entity   
                        if ((postImageRequiredAttendeesCount > preImageRequiredAttendeesCount) && (postImageRequiredAttendeesCount != 0 || preImageRequiredAttendeesCount != 0))
                        {
                            //get records from relationship attendee entity associated to appointment
                            EntityCollection relationshipAttendeeEntityCollection = relationshipAttendee.GetRecords(appointmentId, service);

                            //condition checks required attendees of relationship attendee entity and appointment entity should not be null
                            if (relationshipAttendeeEntityCollection != null && relationshipAttendeeEntityCollection.Entities.Count > 0 &&
                                appointment.Attributes.Contains("requiredattendees") && appointment.Attributes["requiredattendees"] != null)
                            {
                                #region create required attendees object from Relationship Attendee (RA) enity records                         
                                //Add records retrieved from relationship attendee entity to this object
                                RequiredAttendees requiredAttendees = new RequiredAttendees();
                                //This object is required to remove the records from Relationship Attendee (RA) enity to keep RA records in sync with appointment entity required attendee fields
                                RequiredAttendees requiredAttendeesCopy = new RequiredAttendees();
                                requiredAttendees.AddToList(relationshipAttendeeEntityCollection, requiredAttendeesCopy);
                                #endregion create required attendees object from Relationship Attendee (RA) enity records

                                //need to compare attendees retrieved from appointment entity with relationship attendee entity records
                                #region compare attendees retrieved from appointment entity with relationship attendees records and create record if entity record not available in relationship attendee entity or delete record from RA entity if it's not in sync with required attendee field on appointment entity 

                                #region create required attendees object from appointment entity
                                //get required attendees from appointment entity for comparison with Relationship Attendee (RA) records
                                var appointmentRequiredAttendees = (EntityCollection)appointment.Attributes["requiredattendees"];
                                #endregion create required attendees object from appointment entity

                                #region create relationship attendee entity record after comparing required attendees from Relationship Attendee (RA) entity and appointment entity

                                if (appointmentRequiredAttendees != null && appointmentRequiredAttendees.Entities.Count > 0)
                                {
                                    //iterate over appointment required attendees to create Relationship Attendee entity records
                                    foreach (var attendee in appointmentRequiredAttendees.Entities)
                                    {
                                        if (attendee != null && attendee.Contains("partyid") && attendee.Attributes["partyid"] != null)
                                        {
                                            relationshipAttendee.Create(attendee, requiredAttendees, requiredAttendeesCopy, service);
                                        }
                                    }
                                }

                                #endregion

                                #region Delete records from Relationship attendee entity to make it sync with required attendee field
                                //iterate over required Attendees Copy object which contains attendees that are not present in Required Attendes field on Appointment entity
                                relationshipAttendee.Delete(requiredAttendeesCopy, service);
                                #endregion


                                #endregion         
                            }
                        }
                        //condition checks that records are addremoveded from the required attendee field on apponitment entity and therfore we need to delete records in RA entity 
                        else if ((postImageRequiredAttendeesCount < preImageRequiredAttendeesCount) && (postImageRequiredAttendeesCount != 0 || preImageRequiredAttendeesCount != 0))
                        {
                            // here attendes are removed from required attendes list
                            //get the list of objects that are removed - get distict from pre image and post image
                            //iterate over these reocrd and delete it from RA entity

                            //IEnumerable<Entity> entityCollection = preImageRequiredAttendees.Entities.Distinct(postImageRequiredAttendees);

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.ToString());
            }
        }
    }

    public class RelationshipAttendee
    {
        public void AddRelationshipAttendee()
        {

        }
        public EntityCollection GetRecords(Guid appointmentId, IOrganizationService organizationService)
        {
            if (appointmentId != null && appointmentId != Guid.Empty)
            {
                string appointmentIdStr = appointmentId.ToString();
                appointmentIdStr = "F6682D3E-68E1-EB11-BACB-0022486EAE28";
                #region fetch
                string str = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">";
                str += "  <entity name=\"new_relationshipattendees\">";
                str += "    <attribute name=\"new_relationshipattendeesid\" />";
                str += "    <attribute name=\"new_name\" />";
                str += "    <attribute name=\"createdon\" />";
                str += "    <attribute name=\"new_user\" />";
                str += "    <attribute name=\"statuscode\" />";
                str += "    <attribute name=\"statecode\" />";
                str += "    <attribute name=\"overriddencreatedon\" />";
                str += "    <attribute name=\"ownerid\" />";
                str += "    <attribute name=\"modifiedon\" />";
                str += "    <attribute name=\"modifiedonbehalfby\" />";
                str += "    <attribute name=\"modifiedby\" />";
                str += "    <attribute name=\"createdonbehalfby\" />";
                str += "    <attribute name=\"createdby\" />";
                str += "    <attribute name=\"new_contact\" />";
                str += "    <attribute name=\"new_attendeesid\" />";
                str += "    <attribute name=\"new_attendeestatus\" />";
                str += "    <order attribute=\"new_name\" descending=\"false\" />";
                str += "    <filter type=\"and\">";
                str += "      <condition attribute=\"new_attendeesid\" operator=\"eq\" uiname=\"test appointment\" uitype=\"appointment\" value=\"{" + appointmentIdStr + "}\" />";
                str += "    </filter>";
                str += "  </entity>";
                str += "</fetch>";
                #endregion fetch
                //new_relationshipattendees
                return organizationService.RetrieveMultiple(new FetchExpression(str));
            }
            return null;
        }

        public void Create(Entity attendee, RequiredAttendees requiredAttendees, RequiredAttendees requiredAttendeesCopy, IOrganizationService service)
        {
            var entRef = (EntityReference)attendee.Attributes["partyid"];
            if (entRef.LogicalName == "contact")
            {
                Contact contact = new Contact();
                contact.Name = entRef.Name;
                contact.Id = entRef.Id;
                //compare RA required attendee with appointment required attendee
                //if record is not present in RA entity the create a new record
                //else - if recrd is present in RA entity the remove that record from the requiredAttendeesCopy object to keep RA records in sync 
                if (!requiredAttendees.Contacts.Contains(contact))
                {
                    // when required attendee is added to the party list in appointment entity
                    Entity ent = new Entity("new_relationshipattendees");
                    ent.Attributes["new_name"] = entRef.Name;
                    ent.Attributes["new_contact"] = new EntityReference("contact", entRef.Id);
                    ent.Attributes["new_attendeesid"] = new EntityReference("appointment", new Guid("f6682d3e-68e1-eb11-bacb-0022486eae28"));
                    service.Create(ent);
                }
                else
                {
                    requiredAttendeesCopy.Contacts.Remove(contact);
                }
            }
            else if (entRef.LogicalName == "account")
            {
                Account account = new Account();
                account.Name = entRef.Name;
                account.Id = entRef.Id;
                if (!requiredAttendees.Accounts.Contains(account))
                {
                    // when required attendee is added to the party list in appointment entity
                    Entity ent = new Entity("new_relationshipattendees");
                    ent.Attributes["new_name"] = entRef.Name;
                    ent.Attributes["new_account"] = new EntityReference("account", entRef.Id);
                    ent.Attributes["new_attendeesid"] = new EntityReference("appointment", new Guid("f6682d3e-68e1-eb11-bacb-0022486eae28"));
                    service.Create(ent);
                }
                else
                {
                    requiredAttendeesCopy.Accounts.Remove(account);
                }
            }
            else if (entRef.LogicalName == "systemuser")
            {
                User user = new User();
                user.Name = entRef.Name;
                user.Id = entRef.Id;
                if (!requiredAttendees.Users.Contains(user))
                {
                    // when required attendee is added to the party list in appointment entity
                    Entity ent = new Entity("new_relationshipattendees");
                    ent.Attributes["new_name"] = entRef.Name;
                    ent.Attributes["new_user"] = new EntityReference("systemuser", entRef.Id);
                    ent.Attributes["new_attendeesid"] = new EntityReference("appointment", new Guid("f6682d3e-68e1-eb11-bacb-0022486eae28"));
                    service.Create(ent);
                }
                else
                {
                    requiredAttendeesCopy.Users.Remove(user);
                }
            }
        }

        public void Create(EntityCollection attendees, IOrganizationService service)
        {
            if (attendees != null && attendees.Entities.Count > 0)
            {
                foreach (var attendee in attendees.Entities)
                {
                    if (attendee != null && attendee.Contains("partyid") && attendee.Attributes["partyid"] != null)
                    {
                        var entRef = (EntityReference)attendee.Attributes["partyid"];
                        //ent.Id = entRef.Id
                        if (entRef.LogicalName == "contact")
                        {
                            Entity ent = new Entity("new_relationshipattendees");
                            ent.Attributes["new_name"] = entRef.Name;
                            ent.Attributes["new_contact"] = new EntityReference("contact", entRef.Id);
                            ent.Attributes["new_attendeesid"] = new EntityReference("appointment", new Guid("f6682d3e-68e1-eb11-bacb-0022486eae28"));
                            service.Create(ent);
                        }
                        else if (entRef.LogicalName == "account")
                        {
                            Entity ent = new Entity("new_relationshipattendees");
                            ent.Attributes["new_name"] = entRef.Name;
                            ent.Attributes["new_account"] = new EntityReference("account", entRef.Id);
                            ent.Attributes["new_attendeesid"] = new EntityReference("appointment", new Guid("f6682d3e-68e1-eb11-bacb-0022486eae28"));
                            service.Create(ent);

                        }
                        else if (entRef.LogicalName == "systemuser")
                        {
                            Entity ent = new Entity("new_relationshipattendees");
                            ent.Attributes["new_name"] = entRef.Name;
                            ent.Attributes["new_user"] = new EntityReference("systemuser", entRef.Id);
                            ent.Attributes["new_attendeesid"] = new EntityReference("appointment", new Guid("f6682d3e-68e1-eb11-bacb-0022486eae28"));
                            service.Create(ent);
                        }
                    }
                }
            }
        }

        public void Delete(RequiredAttendees requiredAttendeesCopy, IOrganizationService service)
        {
            if (requiredAttendeesCopy.Contacts.Count > 0)
            {
                foreach (var contact in requiredAttendeesCopy.Contacts)
                {
                    service.Delete("contact", contact.Id);
                }
            }

            if (requiredAttendeesCopy.Accounts.Count > 0)
            {
                foreach (var account in requiredAttendeesCopy.Accounts)
                {
                    service.Delete("account", account.Id);
                }
            }


            if (requiredAttendeesCopy.Users.Count > 0)
            {
                foreach (var user in requiredAttendeesCopy.Users)
                {
                    service.Delete("contact", user.Id);
                }
            }
        }

        public void OnCreate()
        {

        }

        public void OnUpdate()
        {

        }
    }
}
