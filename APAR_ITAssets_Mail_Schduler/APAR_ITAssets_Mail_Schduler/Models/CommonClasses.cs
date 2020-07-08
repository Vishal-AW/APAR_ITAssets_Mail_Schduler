using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APAR_ITAssets_Mail_Scheduler.Models
{

    public class ITAssets
    {
        public int Id { get; set; }
        public string RequisitionNo { get; set; }
        public string EmployeeName { get; set; }
        public string Department { get; set; }
        public string Location { get; set; }
        public string Designation { get; set; }
        public string Status { get; set; }
        public string AssignedApprover { get; set; }
        public string Business { get; set; }
        public string StartOn { get; set; }
        public string FunctionalHead { get; set; }
        public string ReportingTo { get; set; }
        public string NewJoineeEmpName  { get; set; }
        public string NewEmployeeDesignation { get; set; }
        public string WorkLevel { get; set; }
        public string ConnectivityDevice { get; set; }
        public string ReplacementEmployeeName { get; set; }
        public string Asset { get; set; }
        public string Modified { get; set; }
        public string StatusId { get; set; }

    }

}
