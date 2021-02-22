using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using MIL.RTI.CourseDocumentGenerator.Constants;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Interfaces;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Abstracts;
using MIL.RTI.CourseDocumentGenerator.Models;
using MIL.RTI.CourseDocumentGenerator.Requests;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel.Updater
{
    public class StudentRecordChecklistHandler : BaseFileUpdater, IUpdateFile
    {
        private const int StartingRowForInserts = 5;
        private const int SoldierNameColumn = 1;
        private const string StartDateCell = "Y3";
        private const string EndDateCell = "AB3";
        private const string CourseCell = "A2";
        private const string InstructorCell = "A3";
        private const string BaseFileName = "Student_Record_Checklist";

        public StudentRecordChecklistHandler(string sourcePath, string targetPath, ClassType classType)
            : base(sourcePath, targetPath, BaseFileName, classType, FileTypes.Excel) { }

        public void UpdateFile(CourseCounselingRequest request)
        {
            var fileName = $"{BaseTargetFileName}_Phase{request.Phase}.xlsx";

            CopyFile(SourcePath, $"{BaseFileName}.xlsx", TargetPath, fileName);

            var fullPath = System.IO.Path.Combine(TargetPath, fileName);

            using (var xlWorkbook = new Workbook(fullPath))
            {
                var workbook = xlWorkbook.OpenWorkbook;

                var worksheet = (Worksheet)workbook.Worksheets.Item[1];
                UpdateInstructor(request.InstructorName, worksheet);
                UpdateDate(StartDateCell, (DateTime)request.CourseStartDate, worksheet);
                UpdateDate(EndDateCell, (DateTime)request.CourseEndDate, worksheet);
                UpdateCourse(request.Course, worksheet);
                AddSoldiers(request.SoldierData, worksheet);
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

        private static void UpdateInstructor(string requestCounselorName, Worksheet worksheet)
        {
            const string cell = InstructorCell;

            var initialValue = worksheet.Range[cell, cell].Value2.ToString();

            var newValue = $"{initialValue} {requestCounselorName}";

            var oRange = worksheet.Range[cell, cell];
            oRange.Cells.Value2 = newValue;
        }

        private static void UpdateDate(string cell, DateTime startDate, _Worksheet worksheet)
        { 
            var initialValue = worksheet.Range[cell, cell].Value2.ToString();

            var newValue = $"{initialValue} {startDate:d}";

            var oRange = worksheet.Range[cell, cell];
            oRange.Cells.Value2 = newValue;
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

                var oRange = worksheet.Range[$"E{currentRow}", $"X{currentRow}"];
                oRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            });
        }
    }
}