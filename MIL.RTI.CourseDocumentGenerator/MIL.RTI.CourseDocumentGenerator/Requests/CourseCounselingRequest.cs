using System.Collections.Generic;
using MIL.RTI.CourseDocumentGenerator.Models;

namespace MIL.RTI.CourseDocumentGenerator.Requests
{
    public class CourseCounselingRequest
    {
        public string CounselorName { get; set; }

        public string Destination { get; set; }

        public string Organization => "2nd Bn 196th RTI";

        public List<SoldierData> SoldierData { get; set; }

        public CounselingData InitialCounseling { get; set; }

        public CounselingData MidCourseCounseling { get; set; }

        public CounselingData EndOfCourseCounseling { get; set; }

        public List<string> Validate()
        {
            var errors = new List<string>();
            
            if (string.IsNullOrEmpty(CounselorName))
            {
                errors.Add("Name And Title of Counselor is Required");
            }

            if (string.IsNullOrEmpty(Destination))
            {
                errors.Add("Destination is Required");
            }

            if (InitialCounseling?.DateOfCounseling == null)
            {
                errors.Add("Initial Counseling Date is Required");
            }

            if (MidCourseCounseling?.DateOfCounseling == null)
            {
                errors.Add("Mid-Course Counseling Date is Required");
            }

            if (EndOfCourseCounseling?.DateOfCounseling == null)
            {
                errors.Add("Initial Counseling Date is Required");
            }

            return errors;
        }
    }
}