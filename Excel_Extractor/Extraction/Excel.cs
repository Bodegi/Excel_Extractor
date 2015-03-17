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

        private static void RowExtract(SheetData sheet, string output, SharedStringTablePart sharedstrings)
        {
            foreach (Row r in sheet.Elements<Row>())
            {
                if(r.RowIndex > 10)
                {
                    List<string> extractedCells = new List<string>();
                    extractedCells = CopyCellValues(r, sharedstrings);
                    InsertCellValues(extractedCells, output);

                }
            }
        }

        private static List<string> CopyCellValues(Row r, SharedStringTablePart sharedstrings)
        {
            List<string> extractedCells = new List<string>();

            foreach (Cell cell in r.Descendants<Cell>())
            {
                if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                {
                    var ssi = sharedstrings.SharedStringTable.Elements<SharedStringItem>().ElementAt(Int32.Parse(cell.CellValue.InnerText));
                    extractedCells.Add(ssi.InnerText);
                }
                else
                {
                    extractedCells.Add(cell.CellValue.InnerText);
                }
            }
            return extractedCells;
        }

        private static void InsertCellValues(List<string> extractedCells, string output)
        {
            using (SpreadsheetDocument FinalFile = SpreadsheetDocument.Open(output, true))
            {
                WorkbookPart wbPart = FinalFile.WorkbookPart;
                WorksheetPart wsPart = wbPart.WorksheetParts.Last();
                SheetData sheetData = wsPart.Worksheet.Elements<SheetData>().First();
                Row row = new Row();
                //Row r = sheetData.Elements<Row>().ElementAt(FinalRowIndex);
                foreach (string extract in extractedCells)
                {
                    Cell cell = new Cell()
                    {
                        CellValue = new CellValue(extract),
                        DataType = CellValues.String

                    };
                    row.Append(cell);
                }
                sheetData.Append(row);
                wsPart.Worksheet.Save();
                wbPart.Workbook.Save();
                FinalFile.Close();
            }
        }

        public static void TabCheck(string file, string output)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(file, true))
            {
                WorkbookPart inputWbPart = document.WorkbookPart;
                int index = 0;
                foreach (WorksheetPart worksheetpart in inputWbPart.WorksheetParts)
                {
                    Worksheet worksheet = worksheetpart.Worksheet;
                    string name = inputWbPart.Workbook.Descendants<Sheet>().ElementAt(index).Name;
                    foreach (SheetData sheetdata in worksheet.Elements<SheetData>())
                    {
                        RowExtract(sheetdata, output, inputWbPart.SharedStringTablePart);
                    }
                    index++;
                }
                //foreach (Sheet sheet in wbPart.Workbook.Descendants<Sheet>())
                //{
                //    //SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                //    string sheetName = sheet.Name;
                //    if(sheetName == "Name")
                //    {
                //        SheetData data = sheet.Elements<SheetData>();
                //        RowExtract(data, FinalFile);
                //    }
                //}
            }
        }
    }
}
