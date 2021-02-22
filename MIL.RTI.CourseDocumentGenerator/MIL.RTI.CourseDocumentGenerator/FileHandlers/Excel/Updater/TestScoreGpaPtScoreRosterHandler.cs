using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using MIL.RTI.CourseDocumentGenerator.Constants;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Interfaces;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Abstracts;
using MIL.RTI.CourseDocumentGenerator.Models;
using MIL.RTI.CourseDocumentGenerator.Requests;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel.Updater
{
    public class TestScoreGpaPtScoreRosterHandler : BaseFileUpdater, IUpdateFile
    {
        private const int StartingRowForInserts = 3;
        private const int SoldierNameColumn = 1;
        private const int RankColumn = 2;
        private const string BaseFileName = "Test_Score_GPA_PT_ScoreRoster";

        public TestScoreGpaPtScoreRosterHandler(string sourcePath, string targetPath, ClassType classType) 
        : base(sourcePath, targetPath, BaseFileName, classType, FileTypes.Excel) {}

        public void UpdateFile(CourseCounselingRequest request)
        {
            var fileName = $"{BaseTargetFileName}_Phase{request.Phase}.xlsx";

            CopyFile(SourcePath, fileName, TargetPath);

            var fullPath = System.IO.Path.Combine(TargetPath, fileName);

            using (var xlWorkbook = new Workbook(fullPath))
            {
                var workbook = xlWorkbook.OpenWorkbook;

                var worksheet = (Worksheet) workbook.Worksheets.Item[1];

                AddSoldiers(request.SoldierData, worksheet);
            }
        }

        private static void AddSoldiers(List<SoldierData> request, _Worksheet worksheet)
        {
            var currentRow = StartingRowForInserts;

            request.ForEach(sd =>
            {
                worksheet.Cells[currentRow, SoldierNameColumn] = sd.FullName;
                worksheet.Cells[currentRow, RankColumn] = sd.Grade;

                currentRow += 1;
            });
            
        }
    }
}