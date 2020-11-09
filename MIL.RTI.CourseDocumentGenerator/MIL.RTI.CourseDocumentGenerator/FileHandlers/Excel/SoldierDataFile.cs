using System.Collections.Generic;
using System.IO;
using System.Linq;
using LinqToExcel;
using MIL.RTI.CourseDocumentGenerator.Models;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel
{
    public class SoldierDataFile
    {
        private static readonly List<string> Columns = new List<string> {"Full Name", "MOS", "Rank/Grade"};
        private readonly ExcelQueryFactory _excel;

        public SoldierDataFile(string path)
        {
            _excel = new ExcelQueryFactory(path);
            Map();
        }

        private void Map()
        {
            _excel.AddMapping<SoldierData>(x => x.FullName, Columns[0]);
            _excel.AddMapping<SoldierData>(x => x.Mos, Columns[1]);
            _excel.AddMapping<SoldierData>(x => x.Rank, Columns[2]);
        }

        public List<SoldierData> GetSoldierData()
        {
            var columns = _excel.GetColumnNames("Sheet1").ToList();

            Columns.ForEach(x =>
            {
                if (columns.All(c => c != x))
                    throw new InvalidDataException(
                        $"Ensure the following Column Headers Exist in Sheet1: {string.Join(", ", Columns)}");
            });

            var worksheet = _excel.Worksheet<SoldierData>();
            var data = worksheet.Select(a => a).ToList();

            data.ForEach(d => d.FullName = d.FullName.Trim());

            return data;
        }
    }
}