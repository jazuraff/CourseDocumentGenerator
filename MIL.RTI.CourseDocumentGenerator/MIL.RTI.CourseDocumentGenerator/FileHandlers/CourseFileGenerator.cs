using MIL.RTI.CourseDocumentGenerator.FileHandlers.Pdf;
using MIL.RTI.CourseDocumentGenerator.Models;
using MIL.RTI.CourseDocumentGenerator.Requests;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers
{
    public class CourseFileGenerator
    {
        private readonly CourseCounselingRequest _request;

        public CourseFileGenerator(CourseCounselingRequest request)
        {
            _request = request;
        }

        public void Execute()
        {
            _request.SoldierData.ForEach(sd =>
            {
                GenerateInitialCounseling(_request.InitialCounseling, sd);
                GenerateMidCourseCounseling(_request.MidCourseCounseling, sd);
                GenerateEndOfCourseCounseling(_request.EndOfCourseCounseling, sd);
            });
        }

        public void GenerateInitialCounseling(CounselingData counselingData, SoldierData soldier)
        {
            var directory = $"{_request.Destination}\\Initial";
            System.IO.Directory.CreateDirectory(directory);
            GenerateDa4856(counselingData, soldier, $"{directory}\\{soldier.FullName}.pdf");
        }

        public void GenerateMidCourseCounseling(CounselingData counselingData, SoldierData soldier)
        {
            var directory = $"{_request.Destination}\\MidCourse";
            System.IO.Directory.CreateDirectory(directory);
            GenerateDa4856(counselingData, soldier, $"{directory}\\{soldier.FullName}.pdf");
        }

        public void GenerateEndOfCourseCounseling(CounselingData counselingData, SoldierData soldier)
        {
            var directory = $"{_request.Destination}\\EndOfCourse";
            System.IO.Directory.CreateDirectory(directory);
            GenerateDa4856(counselingData, soldier, $"{directory}\\{soldier.FullName}.pdf");
        }

        private void GenerateDa4856(CounselingData counselingData, SoldierData soldier, string destination)
        {
            var generator = new Da4856Pdf(".\\Files\\Da4856July2014.pdf", destination);

            generator.GeneratePdf(counselingData, soldier, _request.CounselorName, _request.Organization);
        }
    }
}