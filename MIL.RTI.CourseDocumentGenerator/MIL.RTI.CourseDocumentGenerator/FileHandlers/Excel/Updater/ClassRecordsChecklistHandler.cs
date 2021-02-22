using Microsoft.Office.Interop.Excel;
using MIL.RTI.CourseDocumentGenerator.Constants;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Interfaces;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Abstracts;
using MIL.RTI.CourseDocumentGenerator.Requests;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel.Updater
{
    public class ClassRecordsChecklistHandler : BaseFileUpdater, IUpdateFile
    {
        private const string BaseFileName = "Class_Records_Checklist";
        private const string CourseCell = "A1";
        private const string InstructorCell = "A4";
        private const string ClassCell = "C1";

        public ClassRecordsChecklistHandler(string sourcePath, string targetPath, ClassType classType)
            : base(sourcePath, targetPath, BaseFileName, classType, FileTypes.Excel) { }

        public void UpdateFile(CourseCounselingRequest request)
        {
            var fileName = $"{BaseTargetFileName}_Phase{request.Phase}.xlsx";

            CopyFile(SourcePath, $"{BaseFileName}.xlsx", TargetPath, fileName);

            var fullPath = System.IO.Path.Combine(TargetPath, fileName);

            using (var xlWorkbook = new Workbook(fullPath))
            {
                var workbook = xlWorkbook.OpenWorkbook;

                var worksheet = (Worksheet) workbook.Worksheets.Item[1];

                UpdateClassNumber(request.ClassNumber, worksheet);
                UpdateInstructor(request.InstructorName, worksheet);
                UpdateCourse(request.Course, worksheet);
            }
        }

        private static void UpdateCourse(string requestCourse, Worksheet worksheet)
        {
            const string cell = CourseCell;

            var initialValue = worksheet.Range[cell, cell].Value2.ToString();

            var newValue = $"{initialValue} {requestCourse}";

            var oRange = worksheet.Range[cell, cell];
            oRange.Cells.Value2 = newValue;
        }

        private static void UpdateInstructor(string requestInstructorName, Worksheet worksheet)
        {
            const string cell = InstructorCell;

            var initialValue = worksheet.Range[cell, cell].Value2.ToString();

            var newValue = $"{initialValue} {requestInstructorName}";

            var oRange = worksheet.Range[cell, cell];
            oRange.Cells.Value2 = newValue;
        }

        private static void UpdateClassNumber(string requestClassNumber, _Worksheet worksheet)
        {
            worksheet.Cells[1, 3] = requestClassNumber;
        }
    }
}