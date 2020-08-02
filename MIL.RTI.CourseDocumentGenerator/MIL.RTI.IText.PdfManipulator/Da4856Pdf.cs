using System.Collections.Generic;
using MIL.RTI.PdfDocuments.Constants;
using MIL.RTI.PdfDocuments.Requests;

namespace MIL.RTI.IText.PdfManipulator
{
    public class Da4856Pdf : Pdf
    {
        public Da4856Pdf(string pdfSource, string pdfDestination) : base(pdfSource, pdfDestination) { }

        public void GeneratePdf(CourseCounselingRequest request)
        {
            var fields = new Dictionary<string, string>
            {
                {Da4856July2014Fields.Name, "Still need to do this part"},
                {Da4856July2014Fields.RankGrade, "SSG"},
                {Da4856July2014Fields.DateOfCounseling, request.InitialCounseling.DateOfCounseling?.ToString("ddMMMyyyy")},
                {Da4856July2014Fields.Organization, "Still Need to do this part"},
                {Da4856July2014Fields.NameTitleOfCounselor, request.CounselorName},
                {Da4856July2014Fields.PurposeOfCounseling, request.InitialCounseling.PurposeOfCounseling},
                {Da4856July2014Fields.KeyPointsOfDiscussion, request.InitialCounseling.KeyPoints},
                {Da4856July2014Fields.PlanOfAction, request.InitialCounseling.PlanOfAction},
                {Da4856July2014Fields.LeaderResponsibilities, request.InitialCounseling.LeaderResponsibilities},
                {Da4856July2014Fields.Assessment, request.InitialCounseling.Assessment}
            };

            ManipulateFields(fields);
        }
    }
}