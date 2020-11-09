using System;
using System.Collections.Generic;
using MIL.RTI.CourseDocumentGenerator.Constants;
using MIL.RTI.CourseDocumentGenerator.Constants.CourseDefaults;
using MIL.RTI.CourseDocumentGenerator.Models;

namespace MIL.RTI.CourseDocumentGenerator.Requests
{
    public class CourseCounselingRequest
    {
        public string CounselorName { get; set; }

        public string Destination { get; set; }

        public string Organization => "2nd Bn 196th RTI";

        public string FiscalYear { get; set; }

        public string ClassNumber { get; set; }

        public DateTime? CourseStartDate { get; set; }

        public DateTime? CourseEndDate { get; set; }

        public List<SoldierData> SoldierData { get; set; }

        public CounselingData InitialCounseling { get; set; }

        public CounselingData MidCourseCounseling { get; set; }

        public CounselingData EndOfCourseCounseling { get; set; }

        public string Course
        {
            get
            {
                switch (Class)
                {
                    case ClassType.Mosq:
                        return MosQualificationDefault.CourseNumber;
                    case ClassType.Alc:
                        return AlcDefault.CourseNumber;
                    case ClassType.Slc:
                        return SlcDefault.CourseNumber;
                    default:
                        return "";
                }
            }
        }

        public int Phase { get; set; }

        public ClassType Class { get; set; }

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

            if (FiscalYear == null)
            {
                errors.Add("Fiscal Year is Required");
            }

            if (ClassNumber == null)
            {
                errors.Add("Class Number is Required");
            }

            if (CourseStartDate == null)
            {
                errors.Add("Course Start Date is Required");
            }

            if (CourseEndDate == null)
            {
                errors.Add("Course End Date is Required");
            }

            return errors;
        }
    }
}