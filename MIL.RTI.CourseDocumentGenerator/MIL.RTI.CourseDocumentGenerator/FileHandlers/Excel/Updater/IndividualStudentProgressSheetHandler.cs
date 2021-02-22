using Microsoft.Office.Interop.Excel;
using MIL.RTI.CourseDocumentGenerator.Constants;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Interfaces;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Abstracts;
using MIL.RTI.CourseDocumentGenerator.Requests;
using System;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel.Updater
{
    public class IndividualStudentProgressSheetHandler : BaseFileUpdater, IUpdateFile
    {
        private const string BaseFileName = "Individual_Student_Progress_Sheet";

        public IndividualStudentProgressSheetHandler(string sourcePath, string targetPath, ClassType classType)
            : base(sourcePath, targetPath, BaseFileName, classType, FileTypes.Excel) {}

        public void UpdateFile(CourseCounselingRequest request)
        {
            var fileName = $"{BaseTargetFileName}_Phase{request.Phase}.xlsx";
            
            CopyFile(SourcePath, fileName, TargetPath);

            var fullPath = System.IO.Path.Combine(TargetPath, fileName);

            using (var xlWorkbook = new Workbook(fullPath))
            {
                var workbook = xlWorkbook.OpenWorkbook;

                var worksheet1 = ((Worksheet) workbook.Worksheets[1]);

                //Copy per # of soldiers
                for (var i = 2; i <= request.SoldierData.Count; i++)
                {
                    worksheet1.Copy(Type.Missing, workbook.Worksheets[workbook.Worksheets.Count]);
                    var currentWorkSheet = (Worksheet)workbook.Worksheets[workbook.Worksheets.Count];
                    currentWorkSheet.Name = $"Sheet{i}";
                }

                var currentWorksheet = 1;
                //Update worksheets
                request.SoldierData.ForEach(sd =>
                {
                    var worksheet = (Worksheet)workbook.Worksheets[currentWorksheet];
                    UpdateName(sd.FullName, worksheet);
                    currentWorksheet += 1;
                });
            }
        }

        private static void UpdateName(string name, _Worksheet worksheet)
        {
            const string cell = "A3";

            var initialValue = worksheet.Range[cell, cell].Value2.ToString();

            var newValue = $"{initialValue} {name}";

            worksheet.Cells[3, 1] = newValue;
        }
    }
}