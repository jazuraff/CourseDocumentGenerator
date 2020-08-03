using System.Collections.Generic;
using MIL.RTI.CourseDocumentGenerator.Constants.Fields;
using MIL.RTI.CourseDocumentGenerator.Models;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.Pdf
{
    public class Da4856Pdf : Pdf
    {
        public Da4856Pdf(string pdfSource, string pdfDestination) : base(pdfSource, pdfDestination) { }

        public void GeneratePdf(CounselingData counselingData, SoldierData soldier, string counselorName, string organization)
        {
            var fields = new Dictionary<string, string>
            {
                {Da4856July2014Fields.Name, soldier.FullName},
                {Da4856July2014Fields.RankGrade, soldier.Rank},
                {Da4856July2014Fields.DateOfCounseling, counselingData.DateOfCounseling?.ToString("ddMMMyyyy")},
                {Da4856July2014Fields.Organization, organization},
                {Da4856July2014Fields.NameTitleOfCounselor, counselorName},
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