using System;

namespace MIL.RTI.PdfDocuments.Models
{
    public class CounselingData
    {
        public DateTime? DateOfCounseling { get; set; }

        public string PurposeOfCounseling { get; set; }

        public string KeyPoints { get; set; }

        public string PlanOfAction { get; set; }

        public string LeaderResponsibilities { get; set; }

        public string Assessment { get; set; }
    }
}