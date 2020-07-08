using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APAR_ManpowerRequisition_Mail_Schedular.Models
{

    public class ManpowerRequisition
    {
        public int Id { get; set; }
        public string RequisitionNumber { get; set; }
        public string Author { get; set; }
        public string CreatorDepartment { get; set; }
        public string CreatorLocation { get; set; }
        public string RequisitionDate { get; set; }
        public string Designation { get; set; }
        public string Division { get; set; }
        public string FunctionalHead { get; set; }
        public string HRHead { get; set; }
        public string HRHeadOnly { get; set; }
        public string MDorJMD { get; set; }
        public string Recruiter { get; set; }
        public string Status { get; set; }
        public string IsRejected { get; set; }
        public string CreatedTime { get; set; }
        public string ReplacementEmployeeName { get; set; }
        public string Department { get; set; }
        public string AdditionalBudgets { get; set; }
        //public string DivisionName { get; set; }
        //public string ActiveYear { get; set; }
        public string Modified { get; set; }
        //public string TravelType { get; set; }

    }
    //public class Mailing
    //{
    //    public string MailTo { get; set; }
    //    public string MailCC { get; set; }
    //    public string MailBody { get; set; }
    //    public string MailSubject { get; set; }

    //}


}
