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
        public static int FinalRowIndex { get; set; }

        public Excel()
        {
        }

        private static void RowExtract(Sheet sheet, SpreadsheetDocument FinalFile)
        {
            WorkbookPart wbPart = FinalFile.WorkbookPart;
            WorksheetPart wsPart = wbPart.WorksheetParts.Last();
            SheetData sheetData = wsPart.Worksheet.Elements<SheetData>().First();
            foreach (Row r in sheet.Elements<Row>())
            {
                if(r.RowIndex > 10)
                {
                    sheetData.Append(r);
                    FinalRowIndex = FinalRowIndex + 1;
                }
            }
        }

        public static void TabCheck(string file, SpreadsheetDocument FinalFile)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(file, true))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                foreach (Sheet sheet in wbPart.Workbook.Descendants<Sheet>())
                {
                    string sheetName = sheet.Name;
                    if(sheetName == "Name")
                    {
                        RowExtract(sheet, FinalFile);
                    }
                }
            }
        }
    }
}
