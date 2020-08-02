using MIL.RTI.PdfDocuments.Models;

namespace MIL.RTI.PdfDocuments.Requests
{
    public class CourseCounselingRequest
    {
        public string CounselorName { get; set; }

        public string Destination { get; set; }

        public string SoldierDataFileLocation { get; set; }

        public CounselingData InitialCounseling { get; set; }

        public CounselingData MidCourseCounseling { get; set; }

        public CounselingData EndOfCourseCounseling { get; set; }
    }
}