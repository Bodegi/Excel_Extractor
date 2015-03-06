using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace Extraction
{
    public class Excel
    {
        public Excel()
        {

        }
        public static void TabCheck(string file)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(file, true))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                int count = 0;
                foreach (WorksheetPart worksheetpart in wbPart.WorksheetParts)
                {
                    Worksheet worksheet = worksheetpart.Worksheet;
                    string sheetName = worksheet.Descendants<Sheet>().ElementAt(count).Name;
                    if(sheetName == "Name")
                    {

                    }
                }
            }
        }
    }
}
