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
                if(r.RowIndex > 9)
                {
                    List<Cell> extractedCells = new List<Cell>();
                    extractedCells = CopyCellValues(r, sharedstrings);
                    InsertCellValues(extractedCells, output);

                }
            }
        }

        private static List<Cell> CopyCellValues(Row r, SharedStringTablePart sharedstrings)
        {
            List<Cell> extractedCells = new List<Cell>();

            for (int i = 0; i <= 21; i++)
            {
                Cell cell = r.Descendants<Cell>().ElementAt(i);
                if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                {
                    var ssi = sharedstrings.SharedStringTable.Elements<SharedStringItem>().ElementAt(Int32.Parse(cell.CellValue.InnerText));
                    Cell cellextract = new Cell()
                    {
                        CellValue = new CellValue(ssi.InnerText),
                        DataType = CellValues.String
                    };
                    extractedCells.Add(cellextract);
                }

                else
                {
                    if (cell.CellFormula != null)
                    {
                        int count = cell.CellFormula.Text.Length;
                        string test = cell.InnerText.Substring(cell.InnerText.Length - (cell.InnerText.Length - count));
                        Double cellval = Convert.ToDouble(test);
                        Cell cellextract = new Cell()
                        {
                            CellValue = new CellValue(cellval.ToString()),
                            DataType = CellValues.Number
                        };
                        extractedCells.Add(cellextract);
                    }
                    else
                    {
                        Cell cellextract = new Cell()
                        {
                            CellValue = new CellValue(cell.InnerText),
                            DataType = CellValues.String
                        };
                        extractedCells.Add(cellextract);
                    }
                }
            }

            return extractedCells;
        }

        private static void InsertCellValues(List<Cell> extractedCells, string output)
        {
            using (SpreadsheetDocument FinalFile = SpreadsheetDocument.Open(output, true))
            {
                WorkbookPart wbPart = FinalFile.WorkbookPart;
                WorksheetPart wsPart = wbPart.WorksheetParts.Last();
                SheetData sheetData = wsPart.Worksheet.Elements<SheetData>().First();
                Row row = new Row();
                foreach (Cell extract in extractedCells)
                {
                    row.Append(extract);
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
                foreach (Sheet sheet in inputWbPart.Workbook.Descendants<Sheet>())
                {
                    string name = sheet.Name;
                    if (name.Contains("Income Alloc") == true)
                    {
                        string sheetid = sheet.Id;
                        WorksheetPart wspart = (WorksheetPart)inputWbPart.GetPartById(sheetid);
                        SheetData sdata = wspart.Worksheet.Elements<SheetData>().FirstOrDefault();
                        RowExtract(sdata, output, inputWbPart.SharedStringTablePart);
                    }
                }
            }
        }
    }
}
