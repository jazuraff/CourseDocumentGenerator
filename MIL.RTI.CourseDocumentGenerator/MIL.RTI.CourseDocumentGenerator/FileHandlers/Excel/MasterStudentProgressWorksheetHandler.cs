using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using MIL.RTI.CourseDocumentGenerator.Constants;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel.Interfaces;
using MIL.RTI.CourseDocumentGenerator.Models;
using MIL.RTI.CourseDocumentGenerator.Requests;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel
{
    public class MasterStudentProgressWorksheetHandler : BaseFileUpdater, IUpdateFile
    {
        private const int StartingRowForInserts = 2;
        private const string BaseFileName = "Master_Student_Progress_Worksheet";

        public MasterStudentProgressWorksheetHandler(string sourcePath, string targetPath, ClassType classType)
            : base(sourcePath, targetPath, BaseFileName, classType) { }

        public void UpdateFile(CourseCounselingRequest request)
        {
            var fileName = $"{BaseTargetFileName}_Phase{request.Phase}.xlsx";

            CopyFile(SourcePath, fileName, TargetPath);

            var fullPath = System.IO.Path.Combine(TargetPath, fileName);

            using (var xlWorkbook = new Workbook(fullPath))
            {
                var workbook = xlWorkbook.OpenWorkbook;

                var worksheet = (Worksheet)workbook.Worksheets.Item[1];

                //TODO: this is different depending on the phase and the class
                AddSoldiers(request.SoldierData, worksheet, request.Phase == 1 ? "AE" : "P");
            }
        }

        private static void AddSoldiers(List<SoldierData> request, _Worksheet worksheet, string range)
        {
            var currentRow = StartingRowForInserts;

            request.ForEach(sd =>
            {
                worksheet.Cells[currentRow, 1] = sd.FullName;

                currentRow += 1;
                var line = (Range)worksheet.Rows[currentRow];
                line.Insert();

                var oRange = worksheet.Range[$"B{currentRow}", range+currentRow];
                oRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            });

        }
    }
}