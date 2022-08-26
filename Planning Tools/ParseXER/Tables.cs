using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Planning_Tools.ParseXER
{
    internal class Tables
    {
        public string TableName;
        public string[] Header;
        public List<string[]> Rows;
        public void WriteTableDataToSheet()
        {
            Application xlApp = (Application)ExcelDnaUtil.Application;

            Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null)
                return;

            Worksheet ws = wb.Worksheets.Add(Type.Missing, wb.Worksheets[wb.Worksheets.Count], Type.Missing, Type: XlSheetType.xlWorksheet);
            ws.Name = this.TableName;

            object[,] _headers = new object[1, Header.Length];

            for (int column = 0; column < this.Header.Length; column++)
            {
                _headers[0, column] = this.Header[column];
            }

            object[,] _rows = new object[this.Rows.Count(), Header.Length];
            for (int row = 0; row < this.Rows.Count(); row++)
            {
                for (int column = 0; column < this.Header.Length; column++)
                {
                    _rows[row, column] = this.Rows[row][column];
                }
            }

            ws.Range["A1"].Resize[1, this.Header.Length].Value = _headers;
            ws.Range["A2"].Resize[this.Rows.Count(), this.Header.Length].Value = _rows;
            ws.Range["A1"].Resize[this.Rows.Count() + 1, this.Header.Length].EntireColumn.AutoFit();


            //ws.Columns["A:A"].EntireColumn.AutoFit();
        }
    }
}
