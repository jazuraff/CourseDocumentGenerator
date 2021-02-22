using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using MIL.RTI.CourseDocumentGenerator.Constants;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Interfaces;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Abstracts;
using MIL.RTI.CourseDocumentGenerator.Models;
using MIL.RTI.CourseDocumentGenerator.Requests;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel.Updater
{
    public class PlatoonDutyRosterHandler : BaseFileUpdater, IUpdateFile
    {
        private const int StartingRowForInserts = 2;
        private const int SoldierNameColumn = 1;
        private const string BaseFileName = "PLT_Duty_Roster";

        public PlatoonDutyRosterHandler(string sourcePath, string targetPath, ClassType classType) 
        : base(sourcePath, targetPath, BaseFileName, classType, FileTypes.Excel) {}

        public void UpdateFile(CourseCounselingRequest request)
        {
            var fileName = GetFileName(request.Phase);

            CopyFile(SourcePath, $"{BaseFileName}.xlsx", TargetPath, fileName);

            var fullPath = GetFullPath(request.Phase);

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

                currentRow += 1;
            });
            
        }
    }
}