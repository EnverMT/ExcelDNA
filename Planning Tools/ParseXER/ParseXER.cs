using System;
using System.Linq;
using System.IO;
using System.Collections;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Planning_Tools;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Planning_Tools.ParseXER
{
    interface IParseXER
    {
        void Parse();
    }
    internal class ParseXER : IParseXER
    {
        private List<Tables> tables = new List<Tables>();

        public void Parse()
        {
            string _filepath = Managers.FileManager.GetFilePath();
            if (_filepath == null)
            {
                MessageBox.Show("Incorrect filepath");
                return;
            }

            string[] _fileContent = Managers.FileManager.GetFileContent(_filepath);
            if (_fileContent == null)
            {
                MessageBox.Show("Empty file");
                return;
            }

            ParseFileToTables(_fileContent);

            Application xlApp = (Application)ExcelDnaUtil.Application;

            xlApp.ScreenUpdating = false;
            for (int i = 0; i < tables.Count(); i++)
            {
                xlApp.StatusBar = String.Format($"Processing table {tables[i].TableName}. Total progress {i+1} of {tables.Count()}");
                
                tables[i].WriteTableDataToSheet();                
            }
            xlApp.ScreenUpdating = true;
        }
        private void ParseFileToTables(string[] _fileContent)
        {
            Tables _table = null;
            for (int i = 0; i < _fileContent.Length; i++)
            {
                string[] row = _fileContent[i].Split('\t');
                if (row[0] == "%T")
                {
                    if (_table != null)
                    {
                        tables.Add(_table);
                    }
                    _table = new Tables();
                    _table.TableName = row[1];
                }
                if (row[0] == "%F")
                {
                    _table.Header = row.Skip(1).ToArray();
                    _table.Rows = new List<string[]>();
                }
                if (row[0] == "%R")
                {
                    _table.Rows.Add(row.Skip(1).ToArray()); ;
                }
                if (row[0] == "%E")
                {
                    tables.Add(_table);
                }
            }
        }

    }
}
