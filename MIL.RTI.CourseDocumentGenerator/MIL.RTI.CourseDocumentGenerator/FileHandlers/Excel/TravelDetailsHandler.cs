using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using MIL.RTI.CourseDocumentGenerator.Constants;
using MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel.Interfaces;
using MIL.RTI.CourseDocumentGenerator.Models;
using MIL.RTI.CourseDocumentGenerator.Requests;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel
{
    public class TravelDetailsHandler : BaseFileUpdater, IUpdateFile
    {
        private const string BaseFileName = "Travel_Details";
        private const int StartingRowForSoldierInserts = 7;
        private const int SoldierNameColumn = 1;

        public TravelDetailsHandler(string sourcePath, string targetPath, ClassType classType) 
            : base(sourcePath, targetPath, BaseFileName, classType) {}

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

        private static void AddSoldiers(List<SoldierData> requestSoldierData, Worksheet worksheet)
        {
            var currentRow = StartingRowForSoldierInserts;

            requestSoldierData.ForEach(sd =>
            {
                worksheet.Cells[currentRow, SoldierNameColumn] = sd.FullName;

                currentRow += 1;
            });
        }
    }
}