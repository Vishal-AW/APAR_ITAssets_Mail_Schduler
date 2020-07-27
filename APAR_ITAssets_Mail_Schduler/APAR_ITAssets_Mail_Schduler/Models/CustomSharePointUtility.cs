using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using APAR_ITAssets_Mail_Scheduler.Models;
using UserInformation;
using MSC = Microsoft.SharePoint.Client;
namespace APAR_ITAssets_Mail_Scheduler.Models
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

            //    string LogString = DateTime.Now.ToString("dd/MM/yyyy HH:MM") + " " + logmsg.ToString();

                //  logFile.WriteLine(DateTime.Now);
                //  logFile.WriteLine(logmsg.ToString());
            //    logFile.WriteLine(LogString);

                //logFile.Close();
            }
            catch (Exception ex)
            {
             //   WriteLog(ex.ToString());

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


        public static List<ITAssets> GetAll_AssetsDetailsFromSharePoint(string siteUrl, string listName, string DaysDifference)
        {
            List<ITAssets> _retList = new List<ITAssets>();
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

                            var dataDateValue = DateTime.Now.AddDays(-Convert.ToInt32 (DaysDifference));
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
                                                                <FieldRef Name='StatusId'/>
                                                                <Value Type='Number'>0</Value>
                                                            </Eq>
                                                            <Eq>
                                                                <FieldRef Name='StatusId'/>
                                                                <Value Type='Number'>1</Value>
                                                            </Eq>
                                                        </Or>
                                                        <Or>  
                                                            <Eq>
                                                                <FieldRef Name='StatusId'/>
                                                                <Value Type='Number'>4</Value>
                                                            </Eq>
                                                            <Eq>
                                                                <FieldRef Name='StatusId'/>
                                                                <Value Type='Number'>5</Value>
                                                            </Eq>
                                                        </Or>
                                                    </Or> 
                                                    <Or>  
                                                        <Eq>
                                                            <FieldRef Name='StatusId'/>
                                                            <Value Type='Number'>6</Value>
                                                        </Eq>
                                                        <Eq>
                                                            <FieldRef Name='StatusId'/>
                                                            <Value Type='Number'>7</Value>
                                                        </Eq>
                                                    </Or>
                                            </Or>
                                            <Leq><FieldRef Name='Modified'/><Value Type='DateTime'>" + dataDateValue.ToString("o") + "</Value></Leq>";                                                 
                                            camlQuery.ViewXml += @"</And>
                                    </Where>
                                </Query>
                                <RowLimit>5000</RowLimit>
                                <ViewFields>
                                <FieldRef Name='ID'/>
                                <FieldRef Name='RequisitionNo'/>
                                <FieldRef Name='EmployeeName'/>
                                <FieldRef Name='Department'/>
                                <FieldRef Name='Location'/>
                                <FieldRef Name='Designation'/>
                                <FieldRef Name='Status'/>
                                <FieldRef Name='AssignedApprover'/> 
                                <FieldRef Name='Business'/>
                                <FieldRef Name='StartOn'/>
                                <FieldRef Name='FunctionalHead'/>
                                <FieldRef Name='ReportingTo'/>
                                <FieldRef Name='NewJoineeEmpName'/>
                                <FieldRef Name='WorkLevel'/>
                                <FieldRef Name='ConnectivityDevice'/>
                                <FieldRef Name='ReplacementEmployeeName'/>
                                <FieldRef Name='Asset'/>
                                <FieldRef Name='NewEmployeeDesignation'/>
                                <FieldRef Name='StatusId'/>
                                <FieldRef Name='Modified'/>
                                </ViewFields></View>";
                            MSC.ListItemCollection Items = list.GetItems(camlQuery);

                            context.Load(Items);
                            context.ExecuteQuery();
                            itemPosition = Items.ListItemCollectionPosition;
                            foreach (MSC.ListItem item in Items)
                            {
                                _retList.Add(new ITAssets
                                {
                                    Id = Convert.ToInt32(item["ID"]),
                                    RequisitionNo = Convert.ToString(item["RequisitionNo"]).Trim(),
                                    EmployeeName = Convert.ToString(item["EmployeeName"]).Trim(),
                                    Department = Convert.ToString(item["Department"]).Trim(),
                                    Location = Convert.ToString(item["Location"]).Trim(),
                                    Designation = Convert.ToString(item["Designation"]).Trim(),
                                    Status = Convert.ToString(item["Status"]).Trim(),
                                    StatusId = Convert.ToString(item["StatusId"]).Trim(),
                                    AssignedApprover = item["AssignedApprover"] == null ? "" : Convert.ToString((item["AssignedApprover"] as Microsoft.SharePoint.Client.FieldUserValue[])[0].LookupId),
                                    //Business = item["AssignedApprover"] == null ? "" : Convert.ToString((item["AssignedApprover"] as Microsoft.SharePoint.Client.FieldUserValue[])[0].LookupId),
                                    Business = Convert.ToString(item["Business"]).Trim(),
                                    StartOn = Convert.ToString(item["StartOn"]).Trim(),
                                    FunctionalHead = item["FunctionalHead"] == null ? "" : Convert.ToString((item["FunctionalHead"] as Microsoft.SharePoint.Client.FieldUserValue[])[0].LookupId),
                                    ReportingTo = item["ReportingTo"] == null ? "" : Convert.ToString((item["ReportingTo"] as Microsoft.SharePoint.Client.FieldUserValue).LookupValue),
                                    NewJoineeEmpName = Convert.ToString(item["NewJoineeEmpName"]).Trim(),
                                    ReplacementEmployeeName = Convert.ToString(item["ReplacementEmployeeName"]).Trim(),
                                    NewEmployeeDesignation = Convert.ToString(item["NewEmployeeDesignation"]).Trim(),
                                    Asset = Convert.ToString(item["Asset"]).Trim(),
                                    Modified = Convert.ToString(item["Modified"]).Trim(),
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
                CustomSharePointUtility.WriteLog("Error in  GetAll_AssetsDetailsFromSharePoint()" + " Error:" + ex.Message);
            }
            return _retList;
        }

        public static bool EmailData(List<ITAssets> updationList, string siteUrl, string listName)
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
                                if (updateItem.StatusId == "0")
                                {
                                    _To = updateItem.FunctionalHead;
                                }
                                else
                                {
                                    _To = updateItem.AssignedApprover;
                                }
                                _Subject = "Gentle Reminder"; // + updateItem.ExpVoucherNo + " Travel Voucher Approval is Pending
                                _Body += "Dear User, <br><br>This is to inform you that below request is pending for your Approval.";
                                _Body += "<br><b>Workflow Name :</b> IT Asset Requisition ";
                                _Body += "<br><b>Request No :</b>  " + updateItem.RequisitionNo;
                                _Body += "<br><b>Date of Requisition :</b>  " + updateItem.StartOn;
                                _Body += "<br><b>Requested By  : </b> " + updateItem.EmployeeName;
                                _Body += "<br><b>Location :</b> " + updateItem.Location;
                                _Body += "<br><b>Department :</b> " + updateItem.Department;
                                _Body += "<br><b>Designation :</b> " + updateItem.Designation;
                                _Body += "<br><b>Buisness :</b> " + updateItem.Business;
                                _Body += "<br><b>Reporting to  :</b> " + updateItem.ReportingTo;
                                if(updateItem.Asset == "Desktop") {
                                    _Body += "<br><b>IT Asset Request for  :</b> " + updateItem.Asset;
                                }
                                else
                                {
                                    _Body += "<br><b>Based on Work Level  :</b> " + updateItem.Asset;
                                }
                               
                                _Body += "<br><b>New Joinee Employee :</b> " + updateItem.NewJoineeEmpName;
                                _Body += "<br><b>New Employee designation :</b> " + updateItem.NewEmployeeDesignation;
                                _Body += "<br><b>Replacement Of :</b> " + updateItem.ReplacementEmployeeName;
                                if (updateItem.StatusId == "0")
                                {
                                    _Body += "<br><b>Status :</b> " + updateItem.Status;
                                }
                                else if (updateItem.StatusId == "1")
                                {
                                    _Body += "<br><b>Status :</b> " + updateItem.Status;
                                }
                                else if (updateItem.StatusId == "4")
                                {
                                    _Body += "<br><b>Status :</b> " + updateItem.Status;
                                }
                                else if (updateItem.StatusId == "5")
                                {
                                    _Body += "<br><b>Status :</b> " + updateItem.Status;
                                }
                                else if (updateItem.StatusId == "6")
                                {
                                    _Body += "<br><b>Status :</b> " + updateItem.Status;
                                }
                                else if (updateItem.StatusId == "7")
                                {
                                    _Body += "<br><b>Status :</b> " + updateItem.Status;
                                }
                                _Body += "<br><h3>Kindly provide your approval</h3>";
                                _Body += "<br><h3>For Approval Please Click in the below link</h3>";
                                if (updateItem.StatusId == "0") {
                                    _Body += "<br><a href=\"https://aparindltd.sharepoint.com/ITAssetRequisition/SitePages/PendingForMe.aspx?StatusId="+ updateItem.StatusId + "&RequisitionNo=\">View Link</a>";
                                }
                                else if (updateItem.StatusId == "1") {
                                    _Body += "<br><a href=\"https://aparindltd.sharepoint.com/ITAssetRequisition/SitePages/PendingForMe.aspx?StatusId=" + updateItem.StatusId + "&RequisitionNo=\">View Link</a>";
                                }
                                else if (updateItem.StatusId == "4")
                                {
                                    _Body += "<br><a href=\"https://aparindltd.sharepoint.com/ITAssetRequisition/SitePages/PendingForMe.aspx?StatusId=" + updateItem.StatusId + "&RequisitionNo=\">View Link</a>";
                                }
                                else if (updateItem.StatusId == "5")
                                {
                                    _Body += "<br><a href=\"https://aparindltd.sharepoint.com/ITAssetRequisition/SitePages/PendingForMe.aspx?StatusId=" + updateItem.StatusId + "&RequisitionNo=\">View Link</a>";
                                }
                                else if (updateItem.StatusId == "6")
                                {
                                    _Body += "<br><a href=\"https://aparindltd.sharepoint.com/ITAssetRequisition/SitePages/PendingForMe.aspx?StatusId=" + updateItem.StatusId + "&RequisitionNo=\">View Link</a>";
                                }
                                else if (updateItem.StatusId == "7")
                                {
                                    _Body += "<br><a href=\"https://aparindltd.sharepoint.com/ITAssetRequisition/SitePages/PendingForMe.aspx?StatusId=" + updateItem.StatusId + "&RequisitionNo=\">View Link</a>";
                                }
                                //data.MailTo = _From;
                                //data.MailTo = _To;
                                //data.MailCC = _Cc;
                                //data.MailSubject = _Subject;
                                //data.MailBody = _Body;
                                //varx.Add(data);
                                listItem["ToUser"] = _To;
                                listItem["MailSubject"] = _Subject;
                                listItem["MailBody"] = _Body;
                                //listItem.Update();
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
