using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace MIL.RTI.CourseDocumentGenerator.FileHandlers.Excel
{
    public class Workbook : IDisposable
    {
        private Application _application;
        public Microsoft.Office.Interop.Excel.Workbook OpenWorkbook;
        private List<Worksheet> _worksheets;

        public Workbook(string filePath, int worksheetsCount = 1)
        {
            Open(filePath, worksheetsCount);
        }

        private void Open(string filePath, int workSheetsCount)
        { 
            _application = new Application();
            OpenWorkbook = _application.Workbooks.Open(filePath);

            _worksheets = new List<Worksheet>();

            for (var i = 1; i <= workSheetsCount; i++)
            {
                var xlWorkSheet = (Worksheet) OpenWorkbook.Worksheets.Item[i];
                _worksheets.Add(xlWorkSheet);
            }
        }

        private void CloseWorkbook()
        {
            if (OpenWorkbook == null)
            {
                return;
            }

            var misValue = System.Reflection.Missing.Value;

            OpenWorkbook.Close(true, misValue, misValue);
            
            _application.Quit();

            _worksheets.ForEach(ReleaseObject);

            ReleaseObject(OpenWorkbook);
            ReleaseObject(_application);
        }

        private static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }

        public void Dispose()
        {
            CloseWorkbook();
        }
    }
}