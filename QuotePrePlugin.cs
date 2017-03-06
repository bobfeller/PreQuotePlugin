using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Net.Mail;
using System.Collections;

// Microsoft Dynamics CRM namespace(s)
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using ASG.Crm.Sdk.Web;

namespace ASG.Crm.Sdk.Web
{
    public class QuotePreCreateUpdateHandler : IPlugin
    {
        #region Main Plug-in Execution

        IOrganizationService wService;
        IOrganizationService adminService;
        Guid guidQuote = new Guid();

        /// <summary>
        /// This plugin does a precheck of contact domain name and determines whether account exists, will 
        /// also assign ownership based on account owner.
        /// </summary>
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));
            // Prevent recursion
            if (context.Depth > 1)
                return;

            // Verify we have an entity to work with
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target business entity from the input parmameters.
                Entity entity = (Entity)context.InputParameters["Target"];

                EntityReference cust = new EntityReference();
                EntityReference quoteOwnerOld = new EntityReference();
                EntityReference quoteOwnerNew = new EntityReference();
                string oldEmail = "";
                string newEmail = "";
                string strQuoteName = "";
                bool isUpdate = false;

                // Verify that the entity represents an account.
                if (entity.LogicalName == "quote")
                {
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    wService = serviceFactory.CreateOrganizationService(context.UserId);

                    // est up web service object using admin creds
                    Guid adminUser = FindUser(Properties.Settings.Default.AdminUser, wService, true);
                    adminService = serviceFactory.CreateOrganizationService(adminUser);

                    // See if in Update or Create
                    if (entity.Contains("quoteid"))
                    {
                        if ((Guid)entity["quoteid"] != Guid.Empty)
                        {
                            isUpdate = true;
                            guidQuote = ((Guid)entity["quoteid"]);
                        }
                    }

                    if (entity.Contains("customerid"))
                    {
                        cust = ((EntityReference)entity["customerid"]);
                    }

                    if (isUpdate)
                    {
                        // Look up Quote for Customer 
                        if (!entity.Contains("customerid"))
                        {
                            EntityReference wOwner = new EntityReference();
                            try                            
                            {
                                EntityReference wCust = FindQuote(guidQuote, adminService, ref strQuoteName, ref wOwner);
                                cust.LogicalName = wCust.LogicalName;
                                cust.Id = wCust.Id;
                                quoteOwnerOld.LogicalName = wOwner.LogicalName;
                                quoteOwnerOld.Id = wOwner.Id;
                            }
                            catch { isUpdate = false; }
                        }
                        if (entity.Contains("ownerid"))
                        {
                            quoteOwnerOld.LogicalName = "quote";
                            quoteOwnerOld.Id = ((EntityReference)entity["ownerid"]).Id;
                        }
                    }

                    // If name is in entity use it
                    if(entity.Contains("name"))
                        strQuoteName = entity["name"].ToString();

                    // If new owner specified on form, use it
                    if (entity.Contains("ownerid"))
                        quoteOwnerOld = (EntityReference)entity["ownerid"];

                    if (cust.LogicalName == "account") // Must be an account owner
                    {
                        EntityReference owner = FindAccount(cust.Id, adminService);
                        quoteOwnerNew.LogicalName = owner.LogicalName;
                        quoteOwnerNew.Id = owner.Id;
                    }
                    else    // Must be contact owner
                    {
                        EntityReference owner = FindContact(cust.Id, adminService);
                        quoteOwnerNew.LogicalName = owner.LogicalName;
                        quoteOwnerNew.Id = owner.Id;
                    }

                    // Find email addresses of both the old and new owner
                    if (quoteOwnerOld.Id != Guid.Empty)
                    {
                        oldEmail = FindUser(quoteOwnerOld.Id, adminService);
                    }
                    if (quoteOwnerNew.Id != Guid.Empty)
                    {
                        newEmail = FindUser(quoteOwnerNew.Id, adminService);
                    }

                    // Assign the owner to the quote

                    //if (!String.IsNullOrEmpty(newEmail) || !String.IsNullOrEmpty(oldEmail))
                    //{
                        // compare emails, if different send out email to old owner and new owner
                        //if ((isUpdate) && (newEmail != oldEmail))
                        //{
                        //    string strMsg = "Info: The quote '" + strQuoteName + "' has been re-assigned to " +
                        //        newEmail + " from " + oldEmail + ".\n\n" +
                        //        "You may view this quote at: <" + Properties.Settings.Default.QuoteURL + "{" + guidQuote.ToString() + "}>";
                        //    SendMail(oldEmail, newEmail, "CRM Quote Ownership change", strMsg);
                        //}
                    //}

                }
            }
        }


        #endregion

        #region CRM API Support methods 

        // This method will look up an account based on the Parent Guid
        public EntityReference FindAccount(Guid accountid,IOrganizationService wService)
        {
            try
            {
                QueryExpression query = new QueryExpression();

                // Set the query to retrieve Lead records.
                query.EntityName = "account"; 

                // Create a set of columns to return.
                ColumnSet cols = new ColumnSet(new string[] { "accountid", "ownerid" });
                
                // Create the ConditionExpressions.
                ConditionExpression condition = new ConditionExpression();
                condition.AttributeName = "accountid";
                condition.Operator = ConditionOperator.Equal;
                condition.Values.Add(accountid);

                // Builds the filter based on the condition
                FilterExpression filter = new FilterExpression();
                filter.FilterOperator = LogicalOperator.Or;
                filter.Conditions.Add(condition);

                query.ColumnSet = cols;
                query.Criteria = filter;

                // Retrieve the values from Microsoft CRM.
                EntityCollection retrieved = wService.RetrieveMultiple(query);

                if (retrieved.Entities.Count == 1)
                {
                    Account result = (Account)retrieved.Entities[0];
                    return result.OwnerId;
                }
                else
                {
                    return new EntityReference();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        // This method will look up an account based on the Parent Guid
        public EntityReference FindContact(Guid contactid, IOrganizationService wService)
        {
            try
            {
                QueryExpression query = new QueryExpression();

                // Set the query to retrieve Lead records.
                query.EntityName = "contact";

                // Create a set of columns to return.
                ColumnSet cols = new ColumnSet(new string[] { "contactid", "ownerid" });
                

                // Create the ConditionExpressions.
                ConditionExpression condition = new ConditionExpression();
                condition.AttributeName = "contactid";
                condition.Operator = ConditionOperator.Equal;
                condition.Values.Add(contactid);

                // Builds the filter based on the condition
                FilterExpression filter = new FilterExpression();
                filter.FilterOperator = LogicalOperator.Or;
                filter.Conditions.Add(condition);

                query.ColumnSet = cols;
                query.Criteria = filter;

                // Retrieve the values from Microsoft CRM.
                EntityCollection retrieved = wService.RetrieveMultiple(query);

                if (retrieved.Entities.Count == 1)
                {
                    Contact result = (Contact)retrieved.Entities[0];
                    return result.OwnerId;
                }
                else
                {
                    return new EntityReference();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        // This method will look up a contact and return the Customer, Owner, and Name 
        public EntityReference FindQuote(Guid QuoteId, IOrganizationService wService, ref string QuoteName, ref EntityReference QuoteOwner)
        {
            try
            {
                QueryExpression query = new QueryExpression();

                // Set the query to retrieve Quote record.
                query.EntityName = "quote";

                // Create a set of columns to return.
                ColumnSet cols = new ColumnSet(new string[] { "quoteid", "name", "ownerid", "customerid" });
                
                // Create the ConditionExpressions.
                ConditionExpression condition = new ConditionExpression();
                condition.AttributeName = "quoteid";
                condition.Operator = ConditionOperator.Equal;
                condition.Values.Add(QuoteId);

                // Builds the filter based on the condition
                FilterExpression filter = new FilterExpression();
                filter.FilterOperator = LogicalOperator.And;
                filter.Conditions.Add(condition);

                query.ColumnSet = cols;
                query.Criteria = filter;

                // Retrieve the values from Microsoft CRM.
                EntityCollection retrieved = wService.RetrieveMultiple(query);
                EntityReference owner = new EntityReference();
                if (retrieved.Entities.Count == 1)
                {
                    Quote result = (Quote)retrieved.Entities[0];
                    QuoteName = result.Name;
                    QuoteOwner = result.OwnerId;
                    return result.CustomerId;
                }
                else
                {
                    throw new InvalidPluginExecutionException("Quote not found in CRM database.");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }


        // This method will look up a user based on the Guid, and return email of user if found
        public static string FindUser(Guid guidUser, IOrganizationService wService)
        {
            try
            {
                QueryExpression query = new QueryExpression();

                // Set the query to retrieve User records.
                query.EntityName = "systemuser";

                // Create a set of columns to return.
                ColumnSet cols = new ColumnSet(new string[] { "internalemailaddress", "systemuserid" });

                // Create the ConditionExpressions.
                ConditionExpression condition = new ConditionExpression();
                condition.AttributeName = "systemuserid";
                condition.Operator = ConditionOperator.Equal;
                condition.Values.Add(guidUser);

                // Builds the filter based on the condition
                FilterExpression filter = new FilterExpression();
                filter.FilterOperator = LogicalOperator.And;
                filter.Conditions.Add(condition);

                query.ColumnSet = cols;
                query.Criteria = filter;

                // Retrieve the values from Microsoft CRM.
                EntityCollection retrieved = wService.RetrieveMultiple(query);

                if (retrieved.Entities.Count > 0)
                {
                    SystemUser sysResult = (SystemUser)retrieved.Entities[0];
                    return sysResult.InternalEMailAddress.ToString();
                }
                return "";
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        public static Guid FindUser(string upnsuffix, IOrganizationService wService, bool isUPN)
        {
            try
            {
                QueryExpression query = new QueryExpression();

                // Set the query to retrieve User records.
                query.EntityName = "systemuser";

                // Create a set of columns to return.
                ColumnSet cols = new ColumnSet(new string[] { "internalemailaddress", "systemuserid" });

                // Create the ConditionExpressions.
                ConditionExpression condition = new ConditionExpression();
                condition.AttributeName = "systemuserid";
                condition.Operator = ConditionOperator.Equal;
                condition.Values.Add(upnsuffix);

                // Builds the filter based on the condition
                FilterExpression filter = new FilterExpression();
                filter.FilterOperator = LogicalOperator.And;
                filter.Conditions.Add(condition);

                query.ColumnSet = cols;
                query.Criteria = filter;

                // Retrieve the values from Microsoft CRM.
                EntityCollection retrieved = wService.RetrieveMultiple(query);

                if (retrieved.Entities.Count > 0)
                {
                    SystemUser sysResult = (SystemUser)retrieved.Entities[0];
                    return (Guid)sysResult.SystemUserId;
                }
                return new Guid();
            }
            catch (Exception ex)
            {
                return new Guid();
                //throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        #endregion

        #region Miscellaneous Support functions

        // This method sends out an email for error or notification
        public bool SendMail(string To, string CC, string Subject, string Message)
        {
            try
            {
                if (String.IsNullOrEmpty(To) || String.IsNullOrEmpty(CC)) return true;
                if (To.ToLower() == CC.ToLower()) return true;

                string strSubject = Subject;
                string strBody = Message;

                // Create the Mail Addresses and Mail Message object to send
                MailAddress maFrom = new MailAddress("do_not_reply@virtual.com", "CRM Notification");
                MailAddress maTo = new MailAddress(To);
                MailAddress maCC = new MailAddress(CC);
                //MailAddress maCC2 = new MailAddress("development@virtual.com");
                MailMessage message = new MailMessage(maFrom, maTo);
                message.CC.Add(maCC);
                message.Subject = strSubject;
                message.Body = strBody;
                message.IsBodyHtml = false;
                // Send the mail
                SmtpClient client = new SmtpClient("smtp.virtual.com");
                client.Send(message);
                return true;
            }
            catch (Exception ex)
            {
                // Display the error causing the email to be sent....
                throw new Exception("Unable to send email notification: " + ex.Message);
            }
        }

        // This method will parse the domain name from the email address
        public string GetEmailDomain(string emailaddress)
        {
            string[] strtemp = emailaddress.Split('@');
            if (strtemp.Length < 2)
                throw new InvalidPluginExecutionException("Invalid Email Address: " + emailaddress);
            string[] strtemp2 = strtemp[1].Split('.');
            if (strtemp2.Length < 2)
                throw new InvalidPluginExecutionException("Invalid Email Address: " + emailaddress);

            // Found domain - return
            return strtemp[1];
        }

        #endregion
    }
}
