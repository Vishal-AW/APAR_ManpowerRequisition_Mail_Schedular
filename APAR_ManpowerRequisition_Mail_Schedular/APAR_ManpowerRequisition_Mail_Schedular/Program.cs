using APAR_ManpowerRequisition_Mail_Schedular.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
///using APAR_ManpowerRequisition_Mail_Schedular.Models;

namespace APAR_ManpowerRequisition_Mail_Schedular
{
    class Program
    {
        static void Main()
        {
            //string filename = "log\\Log.txt";
            //CustomSharePointUtility.logFile = new StreamWriter(filename);
            //CustomSharePointUtility.WriteLog("*********************************************");
            //CustomSharePointUtility.WriteLog("Reminder Mail Starts: " + DateTime.Now.ToString());
            //CustomSharePointUtility.WriteLog("*********************************************");
            //Console.WriteLine("*********************************************");
            //Console.WriteLine("Reminder Mail starts : " + DateTime.Now.ToString());
            //Console.WriteLine("*********************************************");
            List<ManpowerRequisition> SPManpowerRequisition = null;
            try
            {
                var siteUrl = ConfigurationManager.AppSettings["SP_Address_Live"];
                string TestManpowerHeaderList = ConfigurationManager.AppSettings["TestManpowerHeaderList"];
                string EmailList = ConfigurationManager.AppSettings["EmailList"];
                string DaysDifference = ConfigurationManager.AppSettings["DaysDifference"];
                //string query = SQLUtility.ReadQuery("EmployeeMasterQuery.txt");
                SPManpowerRequisition = new List<ManpowerRequisition>();
                //Task task_SPEmployeeMaster = Task.Run(() => SPTravelVoucher = CustomSharePointUtility.GetAll_TravelVoucherFromSharePoint(siteUrl, TestingTravelHeaderList));
                SPManpowerRequisition = CustomSharePointUtility.GetAll_ManpowerRequisitionFromSharePoint(siteUrl, TestManpowerHeaderList, DaysDifference);
                //List<TravelVoucher> empMasterFinal = new List<TravelVoucher>();
                List<ManpowerRequisition> empMasterFinal = SPManpowerRequisition;
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
                //CustomSharePointUtility.WriteLog("*********************************************");
                //CustomSharePointUtility.WriteLog("Reminder Mail ends : " + DateTime.Now.ToString());
                //CustomSharePointUtility.WriteLog("*********************************************");
                //Console.WriteLine("*********************************************");
               // Console.WriteLine("Reminder Mail ends : " + DateTime.Now.ToString());
                //Console.WriteLine("*********************************************");
                //CustomSharePointUtility.logFile.Close();
                //Console.ReadKey();

            }
        }
    }
}
