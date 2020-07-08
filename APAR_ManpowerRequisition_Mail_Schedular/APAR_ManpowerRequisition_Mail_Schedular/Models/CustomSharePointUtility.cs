using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using APAR_ManpowerRequisition_Mail_Schedular.Models;
using UserInformation;
using MSC = Microsoft.SharePoint.Client;
namespace APAR_ManpowerRequisition_Mail_Schedular.Models
{
    public static class CustomSharePointUtility
    {
        static UserOperation _UserOperation = new UserOperation();
        public static StreamWriter logFile;
        static byte[] bytes = ASCIIEncoding.ASCII.GetBytes("ZeroCool");
        public static string Decrypt(string cryptedString)
        {
            if (String.IsNullOrEmpty(cryptedString))
            {
                throw new ArgumentNullException("The string which needs to be decrypted can not be null.");
            }

            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream memoryStream = new MemoryStream(Convert.FromBase64String(cryptedString));
            CryptoStream cryptoStream = new CryptoStream(memoryStream, cryptoProvider.CreateDecryptor(bytes, bytes), CryptoStreamMode.Read);
            StreamReader reader = new StreamReader(cryptoStream);

            return reader.ReadToEnd();
        }
        public static MSC.ClientContext GetContext(string siteUrl)
        {
            try
            {
                AppConfiguration _AppConfiguration = GetSharepointCredentials(siteUrl);
                var securePassword = new SecureString();
                foreach (char c in _AppConfiguration.ServicePassword)
                {
                    securePassword.AppendChar(c);
                }
                var onlineCredentials = new MSC.SharePointOnlineCredentials(_AppConfiguration.ServiceUserName, securePassword);
                var context = new MSC.ClientContext(_AppConfiguration.ServiceSiteUrl);
                context.Credentials = onlineCredentials;
                return context;
            }
            catch (Exception ex)
            {
                WriteLog("Error in  CustomSharePointUtility.GetContext: "+ex.ToString());
                return null;
            }
        }
        public static void WriteLog(string logmsg)
        {
            // StreamWriter logFile;

            try
            {

                string LogString = DateTime.Now.ToString("dd/MM/yyyy HH:MM") + " " + logmsg.ToString();

                //  logFile.WriteLine(DateTime.Now);
                //  logFile.WriteLine(logmsg.ToString());
                logFile.WriteLine(LogString);

                //logFile.Close();
            }
            catch (Exception ex)
            {
                WriteLog(ex.ToString());

            }

        }

        public static AppConfiguration GetSharepointCredentials(string siteUrl)
        {
            AppConfiguration _AppConfiguration = new AppConfiguration();

            _AppConfiguration.ServiceSiteUrl = siteUrl;// _UserOperation.ReadValue("SP_Address");
            _AppConfiguration.ServiceUserName = _UserOperation.ReadValue("SP_USER_ID_Live");
            _AppConfiguration.ServicePassword = Decrypt(_UserOperation.ReadValue("SP_Password_Live"));

            return _AppConfiguration;
        }


        public static List<ManpowerRequisition> GetAll_ManpowerRequisitionFromSharePoint(string siteUrl, string listName, string DaysDifference)
        {
            List<ManpowerRequisition> _retList = new List<ManpowerRequisition>();
            try
            {
                using (MSC.ClientContext context = CustomSharePointUtility.GetContext(siteUrl))
                {
                    if (context != null)
                    {
                        MSC.List list = context.Web.Lists.GetByTitle(listName);
                        MSC.ListItemCollectionPosition itemPosition = null;
                        while (true)
                        {
                            var dataDateValue = DateTime.Now.AddDays(-Convert.ToInt32(DaysDifference));
                            MSC.CamlQuery camlQuery = new MSC.CamlQuery();
                            camlQuery.ListItemCollectionPosition = itemPosition;
                            camlQuery.ViewXml = @"<View>
                                 <Query>
                                    <Where>
                                        <And>
                                               <Or> 
                                                    <Or> 
                                                        <Or>  
                                                            <Eq>
                                                                <FieldRef Name='Status'  LookupId='True'/>
                                                                <Value Type='Lookup'>2</Value>
                                                            </Eq>
                                                            <Eq>
                                                                <FieldRef Name='Status'  LookupId='True'/>
                                                                <Value Type='Lookup'>4</Value>
                                                            </Eq>
                                                        </Or>
                                                        <Or>  
                                                            <Eq>
                                                                <FieldRef Name='Status' LookupId='True'/>
                                                                <Value Type='Lookup'>6</Value>
                                                            </Eq>
                                                            <Eq>
                                                                <FieldRef Name='Status'  LookupId='True'/>
                                                                <Value Type='Lookup'>8</Value>
                                                            </Eq>
                                                        </Or>
                                                    </Or> 
                                                <Eq>
                                                    <FieldRef Name='Status'  LookupId='True'/>
                                                    <Value Type='Lookup'>9</Value>
                                                </Eq>
                                            </Or>
                                            <And>
                                            <Eq>
                                                <FieldRef Name='IsRejected'/>
                                                <Value Type='Boolean'>No</Value>
                                            </Eq>
                                            <Leq><FieldRef Name='Modified'/><Value Type='DateTime'>" + dataDateValue.ToString("o") + "</Value></Leq>";


                            camlQuery.ViewXml += @"</And></And></Where></Query>
                                <RowLimit>4000</RowLimit>
                                <ViewFields>
                                <FieldRef Name='ID'/>
                                <FieldRef Name='RequisitionNumber'/>
                                <FieldRef Name='Department'/>
                                <FieldRef Name='CreatorLocation'/>
                                <FieldRef Name='RequisitionDate'/>
                                <FieldRef Name='Designation'/>
                                <FieldRef Name='Division'/>
                                <FieldRef Name='Author'/>
                                <FieldRef Name='FunctionalHead'/>
                                <FieldRef Name='HRHead'/>
                                <FieldRef Name='HRHeadOnly'/>
                                <FieldRef Name='MDorJMD'/>
                                <FieldRef Name='Recruiter'/>
                                <FieldRef Name='Status'/>
                                <FieldRef Name='IsRejected'/>
                                <FieldRef Name='CreatedTime'/>
                                <FieldRef Name='ReplacementEmployeeName'/>
                                <FieldRef Name='AdditionalBudgets'/>
                                <FieldRef Name='Modified'/>
                                </ViewFields></View>";
                            MSC.ListItemCollection Items = list.GetItems(camlQuery);

                            context.Load(Items);
                            context.ExecuteQuery();
                            itemPosition = Items.ListItemCollectionPosition;
                            foreach (MSC.ListItem item in Items)
                            {
                                _retList.Add(new ManpowerRequisition
                                {
                                    Id = Convert.ToInt32(item["ID"]),
                                    RequisitionNumber = Convert.ToString(item["RequisitionNumber"]).Trim(),
                                    Department = Convert.ToString(item["Department"]).Trim(),
                                    CreatorLocation = Convert.ToString(item["CreatorLocation"]).Trim(),
                                    RequisitionDate = Convert.ToString(item["RequisitionDate"]).Trim(),
                                    Designation = Convert.ToString(item["Designation"]).Trim(),
                                    Division = Convert.ToString(item["Division"]).Trim(),
                                    //Author = Convert.ToString((item["Author"] as Microsoft.SharePoint.Client.FieldUserValue).LookupId),
                                    Author = Convert.ToString((item["Author"] as Microsoft.SharePoint.Client.FieldUserValue).LookupValue),
                                    FunctionalHead = Convert.ToString((item["FunctionalHead"] as Microsoft.SharePoint.Client.FieldUserValue[])[0].LookupId),
                                    HRHead = Convert.ToString((item["HRHead"] as Microsoft.SharePoint.Client.FieldUserValue[])[0].LookupId),
                                    HRHeadOnly = Convert.ToString((item["HRHeadOnly"] as Microsoft.SharePoint.Client.FieldUserValue).LookupId),
                                    MDorJMD = item["MDorJMD"] == null ? "" : Convert.ToString((item["MDorJMD"] as Microsoft.SharePoint.Client.FieldUserValue).LookupId),
                                    Recruiter = Convert.ToString((item["Recruiter"] as Microsoft.SharePoint.Client.FieldUserValue[])[0].LookupId),
                                    //Author = Convert.ToString(item["Author"]).Trim(),
                                    //FunctionalHead = Convert.ToString(item["FunctionalHead"]).Trim(),
                                    //HRHead = Convert.ToString(item["HRHead"]).Trim(),
                                    //HRHeadOnly = Convert.ToString(item["HRHeadOnly"]).Trim(),
                                    //MDorJMD = Convert.ToString(item["MDorJMD"]).Trim(),
                                    //Recruiter = Convert.ToString(item["Recruiter"]).Trim(),
                                    //Status = Convert.ToString(item["Status"]).Trim(),
                                    //Status = item["Status"] == null ? "" : Convert.ToString((item["Status"] as Microsoft.SharePoint.Client.FieldUserValue).LookupValue),
                                    Status = Convert.ToString((item["Status"] as Microsoft.SharePoint.Client.FieldLookupValue).LookupValue),
                                    IsRejected = Convert.ToString(item["IsRejected"]).Trim(),
                                    CreatedTime = Convert.ToString(item["CreatedTime"]).Trim(),
                                    ReplacementEmployeeName = Convert.ToString(item["ReplacementEmployeeName"]).Trim(),
                                    AdditionalBudgets = Convert.ToString(item["AdditionalBudgets"]).Trim(),
                                    Modified = Convert.ToString(item["Modified"]),
                                });
                            }
                            if (itemPosition == null)
                            {
                                break; // TODO: might not be correct. Was : Exit While
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CustomSharePointUtility.WriteLog("Error in  GetAll_ManpowerRequisitionFromSharePoint()" + " Error:" + ex.Message);
            }
            return _retList;
        }
        //public static void sample()
        //{
        //    List<TravelVoucher> data = new List<TravelVoucher>();
        //  List<Mailing> mail=EmailData(data, "", "");
        //}

        public static bool EmailData(List<ManpowerRequisition> updationList, string siteUrl, string listName)
        {
            bool retValue = false;
            try
            {

                using (MSC.ClientContext context = CustomSharePointUtility.GetContext(siteUrl))
                {
                    //List<Mailing> varx = new List<Mailing>();

                    MSC.List list = context.Web.Lists.GetByTitle(listName);
                    for (var i = 0; i < updationList.Count; i++)
                    {
                        var updateList = updationList.Skip(i).Take(1).ToList();
                        if (updateList != null && updateList.Count > 0)
                        {
                            foreach (var updateItem in updateList)
                            {
                                MSC.ListItem listItem = null;

                                MSC.ListItemCreationInformation itemCreateInfo = new MSC.ListItemCreationInformation();
                                listItem = list.AddItem(itemCreateInfo);

                                var obj = new Object();
                                //Mailing data = new Mailing();

                                //var _From = "";
                                var _To = "";
                                //var _Cc = "";
                                var _Body = "";
                                var _Subject = "";
                                if (updateItem.Status == "Pending With Functional Head")
                                {
                                    _To = updateItem.FunctionalHead;
                                }
                                else if (updateItem.Status == "Pending With HR Head")
                                {
                                    _To = updateItem.HRHead;
                                }
                                else if (updateItem.Status == "Confirmed By MD And Back to HR Head")
                                {
                                    _To = updateItem.HRHeadOnly;
                                }
                                else if (updateItem.Status == "Pending With MD")
                                {
                                    _To = updateItem.MDorJMD;
                                }
                                else if (updateItem.Status == "Pending With Recruiter")
                                {
                                    _To = updateItem.Recruiter;
                                }
                                _Subject = "Gentle Reminder"; // + updateItem.ExpVoucherNo + " Travel Voucher Approval is Pending
                                _Body += "Dear User, <br><br>This is to inform you that below request is pending for your Approval.";
                                _Body += "<br><b>Workflow Name :</b> Manpower Requisition ";
                                _Body += "<br><b>Voucher No :</b>  " + updateItem.RequisitionNumber;
                                _Body += "<br><b>Date of Creation :</b>  " + updateItem.CreatedTime;
                                _Body += "<br><b>Employee : </b> " + updateItem.Author;
                                _Body += "<br><b>Designation :</b> " + updateItem.Designation;

                                _Body += "<br><b>Department :</b> " + updateItem.Department;
                                _Body += "<br><b>Budgeted / Replacement Additional:</b> " + updateItem.AdditionalBudgets;
                                _Body += "<br><b>Location :</b> " + updateItem.CreatorLocation;
                                _Body += "<br><b>Replacement of :</b> " + updateItem.ReplacementEmployeeName;
                                if (updateItem.Status == "Pending With Functional Head")
                                {
                                    _Body += "<br><b>Status :</b> " + updateItem.Status;
                                }
                                else if (updateItem.Status == "Pending With HR Head")
                                {
                                    _Body += "<br><b>Status :</b> " + updateItem.Status;
                                }
                                else if (updateItem.Status == "Confirmed By MD And Back to HR Head")
                                {
                                    _Body += "<br><b>Status :</b> " + updateItem.Status;
                                }
                                else if (updateItem.Status == "Pending With MD")
                                {
                                    _Body += "<br><b>Status :</b> " + updateItem.Status;
                                }
                                else if (updateItem.Status == "Pending With Recruiter")
                                {
                                    _Body += "<br><b>Status :</b> " + updateItem.Status;
                                }

                                _Body += "<br><h3>Kindly provide your approval</h3>";
                                _Body += "<br><h3>For Approval Please Click in the below link</h3>";
                                if (updateItem.Status == "Pending With Functional Head")
                                {
                                    _Body += "<br><a href=\"https://aparindltd.sharepoint.com/ManpowerRequisition/SitePages/Pending%20With%20Functional%20Head.aspx\">View Link</a>";
                                }
                                else if (updateItem.Status == "Pending With HR Head")
                                {
                                    _Body += "<br><a href=\"https://aparindltd.sharepoint.com/ManpowerRequisition/SitePages/Pending%20With%20HR%20Head.aspx\">View Link</a>";
                                }
                                else if (updateItem.Status == "Confirmed By MD And Back to HR Head")
                                {
                                    _Body += "<br><a href=\"https://aparindltd.sharepoint.com/ManpowerRequisition/SitePages/Pending%20With%20HR%20Head.aspx\">View Link</a>";
                                }
                                else if (updateItem.Status == "Pending With MD")
                                {
                                    _Body += "<br><a href=\"https://aparindltd.sharepoint.com/ManpowerRequisition/SitePages/Pending%20With%20MD.aspx\">View Link</a>";
                                }
                                else if (updateItem.Status == "Pending With Recruiter")
                                {
                                    _Body += "<br><a href=\"https://aparindltd.sharepoint.com/ManpowerRequisition/SitePages/Pending%20With%20Recruiter.aspx\">View Link</a>";
                                }
                               

                                //data.MailTo = _From;
                                //data.MailTo = _To;
                                //data.MailCC = _Cc;
                                //data.MailSubject = _Subject;
                                //data.MailBody = _Body;
                                //varx.Add(data);
                                listItem["ToUser"] = _To;
                                listItem["SubjectDesc"] = _Subject;
                                listItem["BodyDesc"] = _Body;
                                listItem.Update();
                            }
                            try
                            {
                                context.ExecuteQuery();
                                retValue = true;

                            }
                            catch (Exception ex)
                            {
                                CustomSharePointUtility.WriteLog(string.Format("Error in  InsertUpdate_EmployeeMaster ( context.ExecuteQuery();): Error ({0}) ", ex.Message));
                                return false;
                                //continue;
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                CustomSharePointUtility.WriteLog(string.Format("Error in  InsertUpdate_EmployeeMaster: Error ({0}) ", ex.Message));
            }
            return retValue;

        }
    }
    public class AppConfiguration
    {
        public string ServiceSiteUrl;
        public string ServiceUserName;
        public string ServicePassword;
    }
}
