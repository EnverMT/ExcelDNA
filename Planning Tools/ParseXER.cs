using System;
using System.IO;
using System.Collections;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Planning_Tools
{
    static internal class ParseXER
    {
        static public void Parse()
        {
            string _filepath = GetFilePath();
            
            string[] _fileContent = GetFileContent(_filepath);
            if (_fileContent == null) return;

            string[] row = null;
            List<string> headers = new List<string>();
            for (int i = 0; i < _fileContent.Length; i++)
            {
                row = _fileContent[i].Split('\t');
                if (row[0]=="%T")
                {
                    headers.Add(row[1]);
                }
            }
            
        }
        static private string GetFilePath()
        {
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "XER files (*.xer)|*.xer";
                openFileDialog.FilterIndex = 1;                

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {                    
                    filePath = openFileDialog.FileName;
                }
            }
            return filePath;
        }

        static private string[] GetFileContent(string filepath)
        {
            string[] lines = null;
            FileInfo fileInfo = new FileInfo(filepath);
            if (fileInfo.Exists)
            {
                lines = File.ReadAllLines(filepath);
            }
            return lines;
        }
    }
}
