using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace Extraction
{
    public class FileSearch
    {
        public static void traversal(string dir, List<string> visited, SpreadsheetDocument FinalFile)
        {
            bool isExcel = false;
            foreach(string f in Directory.GetFiles(dir))
            {
                isExcel = extension(f);
                if(isExcel == true)
                {
                    Excel.TabCheck(f, FinalFile);
                }
            }
            visited.Add(dir);
            foreach(string d in Directory.GetDirectories(dir))
            {
                if(!visited.Contains(d))
                {
                    traversal(d, visited, FinalFile);
                }
            }
        }

        public static void traversal(string dir, List<string> visited, string output)
        {
            SpreadsheetDocument FinalFile = SpreadsheetDocument.Create(output, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
            bool isExcel = false;
            foreach (string f in Directory.GetFiles(dir))
            {
                isExcel = extension(f);
                if (isExcel == true)
                {
                    Excel.TabCheck(f, FinalFile);
                }
            }
            visited.Add(dir);
            foreach (string d in Directory.GetDirectories(dir))
            {
                if (!visited.Contains(d))
                {
                    traversal(d, visited, FinalFile);
                }
            }
        }

        public static bool extension(string file)
        {
            string ext = "";
            int count = 0;
            count = file.IndexOf(".");
            ext = file.Substring(count).ToUpper();
            if (ext == ".XLS")
            {
                return true;
            }
            if (ext == ".XLSX")
            {
                return true;
            }
            return false;
        }
    }
}
