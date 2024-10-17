using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Reflection.Metadata.Ecma335;

namespace GenPlanilhas
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), @"GeneratedWorkbook.xlsx");
            CreateSpreadsheetWorkbook(filePath);
        }

        public static void CreateSpreadsheetWorkbook(string filepath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                
                workbookPart.Workbook = new Workbook();
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                
                for (int i = 1; i <= 175; i++)
                {
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    Sheet sheet = new Sheet()
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = (uint)i,
                        Name = $"Sheet{i}"
                    };
                    sheets.Append(sheet);

                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                    for (int rowIndex = 1; rowIndex <= 1000; rowIndex++)
                    {
                        Row row = new Row() { RowIndex = (uint)rowIndex };
                        for (int colIndex = 1; colIndex <= 100; colIndex++)
                        {
                            var reference = GetCellReference(colIndex, rowIndex);
                            Cell cell = new Cell()
                            {
                                CellReference = reference,
                                DataType = CellValues.String,
                                CellValue = new CellValue(reference)
                            };


                            // Apply the cell format to the cell
                            cell.StyleIndex = GettRandomStyle(workbookPart);
                            row.Append(cell);
                        }
                        sheetData.Append(row);
                    }
                }

                CreateStyle(workbookPart);

                workbookPart.Workbook.Save();
            }
        }

        private static string GetCellReference(int colIndex, int rowIndex)
        {
            string columnName = GetColumnName(colIndex);
            return $"{columnName}{rowIndex}";
        }

        private static string GetColumnName(int colIndex)
        {
            int dividend = colIndex;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        private static void CreateStyle(WorkbookPart workbookPart)
        {
            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            if (stylesPart.Stylesheet is null)
            {
                stylesPart.Stylesheet = new Stylesheet();
            }

            Stylesheet stylesheet = stylesPart.Stylesheet;

            stylesheet.Fonts = new Fonts(
                new Font(new Color() { Rgb = "FF0000" }), // Red font
                new Font(new Color() { Rgb = "0000FF" }) // Blue font
            );

            stylesheet.Borders = new Borders(
                new Border(
                    new LeftBorder(new Color() { Rgb = "000000" }) { Style = BorderStyleValues.Thin },
                    new RightBorder(new Color() { Rgb = "000000" }) { Style = BorderStyleValues.Thin },
                    new TopBorder(new Color() { Rgb = "000000" }) { Style = BorderStyleValues.Thin },
                    new BottomBorder(new Color() { Rgb = "000000" }) { Style = BorderStyleValues.Thin }
                ),
                new Border(
                    new LeftBorder(new Color() { Rgb = "FF0000" }) { Style = BorderStyleValues.Thick },
                    new RightBorder(new Color() { Rgb = "FF0000" }) { Style = BorderStyleValues.Thick },
                    new TopBorder(new Color() { Rgb = "FF0000" }) { Style = BorderStyleValues.Thick },
                    new BottomBorder(new Color() { Rgb = "FF0000" }) { Style = BorderStyleValues.Thick }
                )
            );

            stylesheet.Fills = new Fills(
                new Fill(new PatternFill() { PatternType = PatternValues.None })
            );

            stylesheet.CellFormats = new CellFormats(
                new CellFormat(), // Default
                new CellFormat() { FontId = 0, BorderId = 0, FillId = 0, ApplyFont = true, ApplyBorder = true }, // Style 1
                new CellFormat() { FontId = 1, BorderId = 0, FillId = 0, ApplyFont = true, ApplyBorder = true }, // Style 2
                new CellFormat() { FontId = 0, BorderId = 1, FillId = 0, ApplyFont = true, ApplyBorder = true }, // Style 3
                new CellFormat() { FontId = 1, BorderId = 1, FillId = 0, ApplyFont = true, ApplyBorder = true }, // Style 4
                new CellFormat() { FontId = 0, BorderId = 0, FillId = 0, ApplyFont = true, ApplyBorder = true }  // Style 5
            );

            stylesheet.Save();
        }

        private static  uint GettRandomStyle(WorkbookPart workbookPart)
        {            
            return (uint)(Random.Shared.Next(5) + 1);
        }        
    }
}