using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Planning_Tools.Managers
{
    static public class FileManager
    {
        static public string GetFilePath()
        {
            string filePath = null;

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

        static public string[] GetFileContent(string filepath)
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
