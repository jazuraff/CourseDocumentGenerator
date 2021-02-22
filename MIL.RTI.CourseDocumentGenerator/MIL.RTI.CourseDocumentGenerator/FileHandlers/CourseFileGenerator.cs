using System;
using System.Collections.Generic;
using MIL.RTI.CourseDocumentGenerator.Constants;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel.Updater;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Interfaces;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Pdf;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Powerpoint.Updater;
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
                
                if (_request.Class == ClassType.Mosq)
                {
                    GenerateMidCourseCounseling(_request.MidCourseCounseling, sd);
                }
               
                GenerateEndOfCourseCounseling(_request.EndOfCourseCounseling, sd);
            });

            GenerateExcelDocs();
        }

        private void GenerateExcelDocs()
        {
            var baseSourcePath = $"{AppDomain.CurrentDomain.BaseDirectory}Files";

            string classSuffix;
            switch (_request.Class)
            {
                case ClassType.Mosq:
                    classSuffix = "13M10";
                    break;
                case ClassType.Alc:
                    classSuffix = "13M30";
                    break;
                case ClassType.Slc:
                    classSuffix = "13M40";
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

            var list = new List<IUpdateFile>
            {
                new TestScoreGpaPtScoreRosterHandler($"{baseSourcePath}/{classSuffix}", _request.Destination, _request.Class), 
                new IndividualStudentProgressSheetHandler($"{baseSourcePath}/{classSuffix}", _request.Destination, _request.Class),
                new MasterStudentProgressWorksheetHandler($"{baseSourcePath}/{classSuffix}", _request.Destination, _request.Class),
                new SignInRosterHandler($"{baseSourcePath}/Shared", _request.Destination, _request.Class),
                new StudentRecordChecklistHandler($"{baseSourcePath}/Shared", _request.Destination, _request.Class),
                new TravelDetailsHandler($"{baseSourcePath}/Shared", _request.Destination, _request.Class),
                new PlatoonDutyRosterHandler($"{baseSourcePath}/Shared", _request.Destination, _request.Class),
                new DaForm87Handler($"{baseSourcePath}/{classSuffix}", _request.Destination, _request.Class),
                new ClassRecordsChecklistHandler($"{baseSourcePath}/Shared", _request.Destination, _request.Class)
            };

            foreach (var fileUpdater in list)
            {
                fileUpdater.UpdateFile(_request);
            }
        }

        private void GenerateInitialCounseling(CounselingData counselingData, SoldierData soldier)
        {
            var directory = $"{_request.Destination}\\Initial";
            System.IO.Directory.CreateDirectory(directory);
            GenerateDa4856(counselingData, soldier, $"{directory}\\{soldier.FullName}.pdf");
        }

        private void GenerateMidCourseCounseling(CounselingData counselingData, SoldierData soldier)
        {
            var directory = $"{_request.Destination}\\MidCourse";
            System.IO.Directory.CreateDirectory(directory);
            GenerateDa4856(counselingData, soldier, $"{directory}\\{soldier.FullName}.pdf");
        }

        private void GenerateEndOfCourseCounseling(CounselingData counselingData, SoldierData soldier)
        {
            var directory = $"{_request.Destination}\\EndOfCourse";
            System.IO.Directory.CreateDirectory(directory);
            GenerateDa4856(counselingData, soldier, $"{directory}\\{soldier.FullName}.pdf");
        }

        private void GenerateDa4856(CounselingData counselingData, SoldierData soldier, string destination)
        {
            var generator = new Da4856Handler(".\\Files\\Shared\\Da4856July2014.pdf", destination);

            generator.GeneratePdf(counselingData, soldier, _request.InstructorName, _request.InstructorTitle, _request.Organization);
        }
    }
}