using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Collections.ObjectModel;
using OfficeDevPnP.Core;
using SP = Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client;
using System.IO;
using System.Xml;
using System.Security.Principal;
using System.Security;
using ControlzEx.Standard;

namespace Alaaliint.Office.Outlook.SubjectSelector
{
    public partial class ThisAddIn
    {
        Microsoft.Office.Interop.Outlook.Inspectors inspectors;

        public ObservableCollection<SubjectTopic> SubjectTopicList = new ObservableCollection<SubjectTopic>();
        public ObservableCollection<SubjectType> SubjectTypeList = new ObservableCollection<SubjectType>();
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;

            inspectors.NewInspector +=
                new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);


            this.Application.ItemSend +=
                new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);

            Reset();
        }

        void Application_ItemSend(object Item, ref bool Cancel)
        {

            Microsoft.Office.Interop.Outlook.MailItem mail = Item as Microsoft.Office.Interop.Outlook.MailItem;
            if (mail != null)
            {
                if (this.CheckUser(mail))
                    this.UpdateEmailStatusTitle(mail.Subject, "Comfirmed");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        // for create ClientId use _layouts/15/AppRegNew.aspx
        // for create permission use _layouts/15/AppInv.aspx
        // for setting _layouts/15/settings.aspx
        // Permission xml

        //<AppPermissionRequests AllowAppOnlyPolicy = "true" >
        //  < AppPermissionRequest Scope="http://sharepoint/content/sitecollection" Right="FullControl"/>
        //  <AppPermissionRequest Scope = "http://sharepoint/content/sitecollection/web" Right="FullControl"/>
        //  <AppPermissionRequest Scope = "http://sharepoint/content/sitecollection/web/list" Right="FullControl"/>
        //</AppPermissionRequests>

        //http://sharepoint/content/sitecollection
        //http://sharepoint/content/sitecollection/web
        //http://sharepoint/content/sitecollection/web/list
        //http://sharepoint/content/tenant

        //Read
        //Write
        //Manage
        //FullControl
        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Microsoft.Office.Interop.Outlook.MailItem mailItem = Inspector.CurrentItem as Microsoft.Office.Interop.Outlook.MailItem;

            ShowSubjectForm(mailItem);
        }

        public void ShowSubjectForm(Microsoft.Office.Interop.Outlook.MailItem mailItem)
        {
            try
            {

                if (mailItem != null)
                {
                    if (mailItem.EntryID == null)
                    {
                        if (CheckUser(mailItem))
                        {
                            frmMain frm = new frmMain();
                            frm.myMailItem = mailItem;
                            if (frm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                if (frm.Subject != null)
                                {
                                    mailItem.Subject = frm.Subject;
                                }

                            }

                            frm.Dispose();
                        }
                        else
                        {
                            System.Windows.Forms.MessageBox.Show("User not found", "AL-AALI Email Template", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
        }
        public void Reset()
        {
            try
            {
                getEmailSubjectFromSharePointOnline();

                getEmailSubjectTypesFromSharePointOnline();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }
        }

        public void getEmailSubjectFromSharePointOnline()
        {
            try
            {
                string siteUrl = Properties.Settings.Default.ClientContextUrl;
                using (var clientContext = new AuthenticationManager().GetAppOnlyAuthenticatedContext(siteUrl, Properties.Settings.Default.ClientId, Properties.Settings.Default.ClientSecret))
                {
                    Web oWebsite = clientContext.Web;
                    ListCollection collList = oWebsite.Lists;
                    SP.List oList = collList.GetByTitle("EmailSubject_Topics");
                    ListItemCollection collListItem = oList.GetItems(new CamlQuery());

                    clientContext.Load(collListItem,
                        items => items.Include(
                        item => item.Id,
                        item => item["Title"],
                        item => item["Code"]));


                    clientContext.ExecuteQuery();
                    Globals.ThisAddIn.SubjectTopicList.Clear();

                    foreach (ListItem oListItem in collListItem)
                    {
                        SubjectTopicList.Add(new SubjectTopic()
                        {
                            ID = oListItem.Id,
                            Title = oListItem["Title"].ToString(),
                            Code = oListItem["Code"].ToString()
                        });

                    }
                };
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }

        }
        public void getEmailSubjectTypesFromSharePointOnline()
        {
            try
            {
                string siteUrl = Properties.Settings.Default.ClientContextUrl;
                using (var clientContext = new AuthenticationManager().GetAppOnlyAuthenticatedContext(siteUrl, Properties.Settings.Default.ClientId, Properties.Settings.Default.ClientSecret))
                {
                    Web oWebsite = clientContext.Web;
                    ListCollection collList = oWebsite.Lists;
                    SP.List oList = collList.GetByTitle("EmailSubject_Types");
                    ListItemCollection collListItem = oList.GetItems(new CamlQuery());

                    clientContext.Load(collListItem,
                        items => items.Include(
                        item => item.Id,
                        item => item["Title"],
                        item => item["Parent"],
                        item => item["Code"]));


                    clientContext.ExecuteQuery();
                    Globals.ThisAddIn.SubjectTypeList.Clear();

                    foreach (ListItem oListItem in collListItem)
                    {
                        string codeValue = "";
                        var s = GetSubjectTopicById(((Microsoft.SharePoint.Client.FieldLookupValue)oListItem["Parent"]).LookupId);
                        if (s != null)
                            codeValue = s.Code;
                        SubjectTypeList.Add(new SubjectType()
                        {

                            ID = oListItem.Id,
                            Title = oListItem["Title"].ToString(),
                            Code = oListItem["Code"].ToString(),
                            Parent = new SubjectTopic(
                                ((Microsoft.SharePoint.Client.FieldLookupValue)oListItem["Parent"]).LookupId,
                                codeValue,
                              ((Microsoft.SharePoint.Client.FieldLookupValue)oListItem["Parent"]).LookupValue.ToString())
                        });

                    }
                };
            }
            catch (Exception ex)
            {
                throw new Exception(ex.ToString());
            }

        }

        private Microsoft.Office.Interop.Outlook.ExchangeUser GetCurrentUserInfo(Microsoft.Office.Interop.Outlook.MailItem mailItem)
        {
            try
            {


                Microsoft.Office.Interop.Outlook.AddressEntry addrEntry =
                    mailItem.SendUsingAccount.CurrentUser.AddressEntry;

                if (addrEntry.Type == "EX")
                {
                    Microsoft.Office.Interop.Outlook.ExchangeUser currentUser =
                        mailItem.SendUsingAccount.CurrentUser.AddressEntry.GetExchangeUser();

                    if (currentUser != null)
                    {
                        return currentUser;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                return null;
                throw new Exception(ex.ToString());
            }

        }

        private bool CheckUser(Microsoft.Office.Interop.Outlook.MailItem mailItem)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.ExchangeUser currentUser = GetCurrentUserInfo(mailItem);
                if (currentUser != null)
                {
                    string siteUrl = Properties.Settings.Default.ClientContextUrl;

                    using (var clientContext = new AuthenticationManager().GetAppOnlyAuthenticatedContext(siteUrl, Properties.Settings.Default.ClientId, Properties.Settings.Default.ClientSecret))
                    {

                        GroupCollection collGroup = clientContext.Web.SiteGroups;

                        clientContext.Load(collGroup);

                        clientContext.Load(collGroup,
                 groups => groups.Include(
                     group => group.Title,
                     group => group.Id,
                     group => group.Users.Include(
                         user => user.Title,
                         user => user.Email)));

                        clientContext.ExecuteQuery();

                        foreach (Group oGroup in collGroup)
                        {

                            UserCollection collUser = oGroup.Users;

                            foreach (User oUser in collUser)
                            {
                                if (currentUser.PrimarySmtpAddress == oUser.Email)
                                {
                                    return true;
                                }
                            }

                        }
                    }
                }


                return false;
            }
            catch (Exception ex)
            {
                return false;
                throw new Exception(ex.ToString());
            }
        }

        private User getEnsureUser(Microsoft.Office.Interop.Outlook.MailItem mailItem)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.ExchangeUser currentUser = GetCurrentUserInfo(mailItem);
                if (currentUser != null)
                {

                    string siteUrl = Properties.Settings.Default.ClientContextUrl;
                    using (var clientContext = new AuthenticationManager().GetAppOnlyAuthenticatedContext(siteUrl, Properties.Settings.Default.ClientId, Properties.Settings.Default.ClientSecret))
                    {
                        var rootWeb = clientContext.Site.RootWeb;
                        var usr = rootWeb.EnsureUser(currentUser.PrimarySmtpAddress);
                        clientContext.Load(usr);
                        clientContext.ExecuteQuery();

                        if (usr != null)
                            return usr;
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                return null;
                throw new Exception(ex.ToString());
            }
        }

        private bool CheckEnsureUser(Microsoft.Office.Interop.Outlook.MailItem mailItem)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.ExchangeUser currentUser = GetCurrentUserInfo(mailItem);
                if (currentUser != null)
                {

                    string siteUrl = Properties.Settings.Default.ClientContextUrl;
                    using (var clientContext = new AuthenticationManager().GetAppOnlyAuthenticatedContext(siteUrl, Properties.Settings.Default.ClientId, Properties.Settings.Default.ClientSecret))
                    {
                        var rootWeb = clientContext.Site.RootWeb;
                        var usr = rootWeb.EnsureUser(currentUser.PrimarySmtpAddress);
                        clientContext.Load(usr);
                        clientContext.ExecuteQuery();

                        if (usr != null)
                            return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                return false;
                throw new Exception(ex.ToString());
            }
        }
        public string InsertEmailSubjectSerial(string subject, string status, SubjectTopic subjectTopicValue, SubjectType subjectTypeValue, Microsoft.Office.Interop.Outlook.MailItem mailItem)
        {
            try
            {
                Guid g = Guid.NewGuid();
                string siteUrl = Properties.Settings.Default.ClientContextUrl;
                using (var clientContext = new AuthenticationManager().GetAppOnlyAuthenticatedContext(siteUrl, Properties.Settings.Default.ClientId, Properties.Settings.Default.ClientSecret))
                {
                    Web oWebsite = clientContext.Web;
                    ListCollection collList = oWebsite.Lists;

                    SP.List oList = collList.GetByTitle("EmailSubject_Serial");
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

                    ListItem oListItem = oList.AddItem(itemCreateInfo);

                    // type section//////////////////////////////////////
                    FieldLookupValue typeFieldLookupValue = new FieldLookupValue();  //GetLookupFieldValue(subjectTypeValue.Title, "EmailSubject_Types");
                    typeFieldLookupValue.LookupId = subjectTypeValue.ID;
                    if (typeFieldLookupValue != null)
                    {
                        oListItem["ParentType"] = typeFieldLookupValue;
                    }
                    ////////////////////////////////////////////////////////////////
                    // tpoic section//////////////////////////////////////
                    FieldLookupValue topicFieldLookupValue = new FieldLookupValue();
                    topicFieldLookupValue.LookupId = subjectTopicValue.ID;
                    //FieldLookupValue topicFieldLookupValue = GetLookupFieldValue(subjectTopicValue.Title, "EmailSubject_Topics");
                    if (topicFieldLookupValue != null)
                    {
                        oListItem["ParentTopic"] = topicFieldLookupValue;
                    }
                    ////////////////////////////////////////////////////////////////
                    // status section//////////////////////////////////////
                    oListItem["Status"] = status; //Pending Comfirmed
                    ///////////////////////////////////////////////////////
                    // user section//////////////////////////////////////////
                    var usr = getEnsureUser(mailItem);
                    if (usr == null)
                        return "";

                    Microsoft.SharePoint.Client.FieldUserValue _userValue = new Microsoft.SharePoint.Client.FieldUserValue();
                    _userValue.LookupId = usr.Id;
                    oListItem["User"] = _userValue;
                    ///////////////////////////////////////////////////////////////
                    // title section///////////////////////////////////////////////
                    oListItem["Title"] = g.ToString();
                    ////////////////////////////////////////////////////////////////
                    oListItem.Update();

                    clientContext.ExecuteQuery();

                    int id = oListItem.Id;

                    string subjectToReturn = "";
                    oListItem["Title"] = string.Format("[{0:D}-{1:D}-{2:D}] {3:D}", subjectTopicValue.Code, subjectTypeValue.Code, id, subject);
                    subjectToReturn = string.Format("[{0:D}-{1:D}-{2:D}] {3:D}", subjectTopicValue.Code, subjectTypeValue.Code, id, subject);

                    oListItem.Update();

                    clientContext.ExecuteQuery();

                    return subjectToReturn;
                }
            }
            catch (Exception ex)
            {
                return "";
                throw new Exception(ex.ToString());
            }
        }

        public static FieldLookupValue GetLookupFieldValue(string lookupName, string lookupListName)
        {
            string siteUrl = Properties.Settings.Default.ClientContextUrl;
            using (var clientContext = new AuthenticationManager().GetAppOnlyAuthenticatedContext(siteUrl, Properties.Settings.Default.ClientId, Properties.Settings.Default.ClientSecret))
            {
                var lookupList = clientContext.Web.Lists.GetByTitle(lookupListName);
                CamlQuery query = new CamlQuery();
                string lookupFieldName = "Title";
                string lookupFieldType = "Text";

                query.ViewXml = string.Format(@"<View><Query><Where><Eq><FieldRef Name='{0}'/><Value Type='{1}'>{2}</Value></Eq>" +
                                                "</Where></Query></View>", lookupFieldName, lookupFieldType, lookupName);

                ListItemCollection listItems = lookupList.GetItems(query);
                clientContext.Load(listItems, items => items.Include
                                                    (listItem => listItem["ID"],
                                                    listItem => listItem[lookupFieldName]));
                clientContext.ExecuteQuery();

                if (listItems != null)
                {
                    ListItem item = listItems[0];
                    FieldLookupValue lookupValue = new FieldLookupValue();
                    lookupValue.LookupId = int.Parse(item["ID"].ToString());
                    return lookupValue;
                }

                return null;
            }

        }
        private void UpdateEmailStatusTitle(string title, string status)
        {
            string siteUrl = Properties.Settings.Default.ClientContextUrl;
            using (var clientContext = new AuthenticationManager().GetAppOnlyAuthenticatedContext(siteUrl, Properties.Settings.Default.ClientId, Properties.Settings.Default.ClientSecret))
            {
                SP.List oListUpdate = clientContext.Web.Lists.GetByTitle("EmailSubject_Serial");
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = string.Format(@"<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>{0}</Value></Eq></Where></Query><ViewFields><FieldRef Name='ID'/><FieldRef Name='Title'/><FieldRef Name='Status'/></ViewFields><RowLimit>1</RowLimit></View>", title);

                ListItemCollection collListItem = oListUpdate.GetItems(camlQuery);
                clientContext.Load(collListItem, items => items.Include(item => item["ID"], item => item["Title"], item => item["Status"]));
                clientContext.ExecuteQuery();

                if (collListItem != null && collListItem.Count > 0)
                {
                    ListItem item = collListItem[0];
                    item["Status"] = status;
                    item.Update();
                    clientContext.ExecuteQuery();
                }


            }

        }

        private SubjectTopic GetSubjectTopicById(int id)
        {
            foreach (SubjectTopic item in SubjectTopicList)
            {
                if (id == item.ID)
                {
                    return item;
                }
            }
            return null;
        }

        private SubjectType GetSubjectTypeById(int id)
        {
            foreach (SubjectType item in SubjectTypeList)
            {
                if (id == item.ID)
                {
                    return item;
                }
            }
            return null;
        }
    }
}
