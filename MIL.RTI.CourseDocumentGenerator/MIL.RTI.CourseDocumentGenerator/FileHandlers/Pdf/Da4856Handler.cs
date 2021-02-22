using System.Collections.Generic;
using MIL.RTI.CourseDocumentGenerator.Constants.Fields;
using MIL.RTI.CourseDocumentGenerator.Helper;
using MIL.RTI.CourseDocumentGenerator.Models;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.Pdf
{
    public class Da4856Handler : Pdf
    {
        public Da4856Handler(string pdfSource, string pdfDestination) : base(pdfSource, pdfDestination) { }

        public void GeneratePdf(CounselingData counselingData, SoldierData soldier, string instructorName, string instructorTitle, string organization)
        {
            var fields = new Dictionary<string, string>
            {
                {Da4856July2014Fields.Name, soldier.FullName},
                {Da4856July2014Fields.RankGrade, $"{soldier.Grade.ToRank()}/{soldier.Grade}"},
                {Da4856July2014Fields.DateOfCounseling, counselingData.DateOfCounseling?.ToString("ddMMMyyyy")},
                {Da4856July2014Fields.Organization, organization},
                {Da4856July2014Fields.NameTitleOfCounselor, $"{instructorName}, {instructorTitle}"},
                {Da4856July2014Fields.PurposeOfCounseling, counselingData.PurposeOfCounseling},
                {Da4856July2014Fields.KeyPointsOfDiscussion, counselingData.KeyPoints},
                {Da4856July2014Fields.PlanOfAction, counselingData.PlanOfAction},
                {Da4856July2014Fields.LeaderResponsibilities, counselingData.LeaderResponsibilities},
                {Da4856July2014Fields.Assessment, counselingData.Assessment}
            };

            ManipulateFields(fields);
        }
    }
}