using APAR_ITAssets_Mail_Scheduler.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
///using APAR_ManpowerRequisition_Mail_Schedular.Models;

namespace APAR_ITAssets_Mail_Scheduler
{
    class Program
    {
        static void Main()
        {

            List<ITAssets> SPITAssets = null;
            try
            {
                var siteUrl = ConfigurationManager.AppSettings["SP_Address_Live"];
                string TestAssetsRequisition = ConfigurationManager.AppSettings["TestAssetsRequisition"];
                string EmailList = ConfigurationManager.AppSettings["EmailList"];
                string DaysDifference = ConfigurationManager.AppSettings["DaysDifference"];
                //string query = SQLUtility.ReadQuery("EmployeeMasterQuery.txt");
                SPITAssets = new List<ITAssets>();
                //Task task_SPEmployeeMaster = Task.Run(() => SPTravelVoucher = CustomSharePointUtility.GetAll_TravelVoucherFromSharePoint(siteUrl, TestingTravelHeaderList));
                SPITAssets = CustomSharePointUtility.GetAll_AssetsDetailsFromSharePoint(siteUrl, TestAssetsRequisition, DaysDifference);
                //List<TravelVoucher> empMasterFinal = new List<TravelVoucher>();
                List<ITAssets> empMasterFinal = SPITAssets;
                if (empMasterFinal.Count > 0)
                {
                    //Console.WriteLine("Employee data synchronized successfully.");
                    var success = CustomSharePointUtility.EmailData(empMasterFinal, siteUrl, EmailList);
                    if (success)
                    {
                        ///CustomSharePointUtility.WriteLog("Reminder Mail Sent Successfully.");
                        //Console.WriteLine("Reminder Mail Sent Successfully.");
                    }
                }
                else
                {
                    //CustomSharePointUtility.WriteLog("No Pending Records.");
                    //Console.WriteLine("No Pending Records.");
                }
            }
            catch (Exception ex)
            {
                CustomSharePointUtility.WriteLog("Error in scheduler : " + ex.StackTrace);
                Console.WriteLine("Error in scheduler : " + ex.StackTrace);
            }
            finally
            {

            }
        }
    }
}
