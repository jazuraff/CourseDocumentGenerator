using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using LinqToExcel;
using MIL.RTI.CourseDocumentGenerator.Models;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel
{
    public class SoldierDataFileHandler : IDisposable
    {
        private const string Name = "Name";
        private const string Mos = "PMOSEN";
        private const string Grade = "Pay Grade";

        private static readonly List<string> Columns = new List<string> {Name, Mos, Grade};
        private readonly ExcelQueryFactory _excel;

        public SoldierDataFileHandler(string path)
        {
            _excel = new ExcelQueryFactory(path);
        }

        private void Map(string columnName)
        {
            switch (columnName)
            {
                case Name:
                    _excel.AddMapping<SoldierData>(sd => sd.FullName, columnName);
                    break;
                case Mos:
                    _excel.AddMapping<SoldierData>(sd => sd.Mos, columnName);
                    break;
                case Grade:
                    _excel.AddMapping<SoldierData>(sd => sd.Grade, columnName);
                    break;
            }
        }

        public List<SoldierData> GetSoldierData()
        {
            var columns = _excel.GetColumnNames("Sheet1").ToList();

            if (columns[0].Contains("FOR OFFICIAL USE ONLY"))
            {
                throw new InvalidDataException(
                    $"Please remove the header data from the ATRRS roster to proceed");
            }

            Columns.ForEach(x =>
            {
                if (columns.All(c => c != x))
                    throw new InvalidDataException(
                        $"Ensure the following Column Headers Exist in Sheet1: {string.Join(", ", Columns)}");

                Map(x);
            });

            var worksheet = _excel.Worksheet<SoldierData>();
            var data = worksheet.Select(a => a).ToList();

            data.ForEach(d => d.FullName = d.FullName.Trim());
                       
            return data;
        }

        public void Dispose()
        {
            _excel.Dispose();
        }
    }
}