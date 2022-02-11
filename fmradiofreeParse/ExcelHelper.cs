using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace fmradiofreeParse
{
    class ExcelHelper : IDisposable
    {
        private Application _excel;
        private Workbook _workbook;
        private string _filePath;

        public ExcelHelper()
        {
            _excel = new Excel.Application();
        }

        public void Dispose()
        {
            _workbook.Close();
            _excel.Quit();
        }

        internal bool Open(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    _workbook = _excel.Workbooks.Open(filePath);
                }
                else
                {
                    _workbook = _excel.Workbooks.Add();
                    _filePath = filePath;
                }
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }

        internal bool Set(int colum, int row, string data)
        {
            try
            {
                ((Worksheet)_excel.ActiveSheet).Cells[colum, row] = data;
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }

        internal void Save()
        {
            if (!string.IsNullOrEmpty(_filePath))
            {
                _workbook.SaveAs(_filePath);
            }
            else _workbook.Save();
        }
    }
}
