using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using MIL.RTI.CourseDocumentGenerator.Constants;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Interfaces;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Abstracts;
using MIL.RTI.CourseDocumentGenerator.Models;
using MIL.RTI.CourseDocumentGenerator.Requests;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel.Updater
{
    public class SignInRosterHandler : BaseFileUpdater, IUpdateFile
    {
        private const string BaseFileName = "Sign_In_Roster";
        private const int StartingRowForInserts = 5;
        private const int SoldierNameColumn = 1;

        public SignInRosterHandler(string sourcePath, string targetPath, ClassType classType)
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
                AddSoldiers(request.SoldierData, worksheet);
            }
        }

        private static void UpdateCourse(string requestCourse, Worksheet worksheet)
        {
            worksheet.Cells[2, 6] = requestCourse;
        }

        private static void UpdateInstructor(string requestCounselorName, Worksheet worksheet)
        {
            worksheet.Cells[1, 6] = requestCounselorName;
        }

        private static void UpdateClassNumber(string requestClassNumber, _Worksheet worksheet)
        {
            const string cell = "A1";

            var initialValue = worksheet.Range[cell, cell].Value2.ToString();

            var newValue = $"{initialValue} {requestClassNumber}";

            worksheet.Cells[1, 1] = newValue;
        }

        private static void AddSoldiers(List<SoldierData> request, _Worksheet worksheet)
        {
            var currentRow = StartingRowForInserts;

            request.ForEach(sd =>
            {
                worksheet.Cells[currentRow, SoldierNameColumn] = sd.FullName;

                currentRow += 1;
                var line = (Range)worksheet.Rows[currentRow];
                line.Insert();
            });
            
        }
    }
}