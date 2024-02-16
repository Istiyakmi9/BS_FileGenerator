using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

namespace BS_FileGenerator.Service
{
    public class DataTableToExcel
    {
        public void ToExcel(DataTable table, string filepath, string sheetName = null)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);
                Sheets? sheets = spreadsheetDocument.WorkbookPart!.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet
                {
                    Id = (StringValue)spreadsheetDocument.WorkbookPart!.GetIdOfPart(worksheetPart),
                    SheetId = (UInt32Value)1u,
                    Name = (StringValue)((sheetName == null) ? "mySheet" : sheetName)
                };
                sheets!.Append(sheet);
                BuildExcelRows(sheetData, table);
                workbookPart.Workbook.Save();
            }
        }

        public void ToExcelWithHeaderAnddata(List<string> header, dynamic data, string filepath, string sheetName = null)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);
                Sheets? sheets = spreadsheetDocument.WorkbookPart!.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet
                {
                    Id = (StringValue)spreadsheetDocument.WorkbookPart!.GetIdOfPart(worksheetPart),
                    SheetId = (UInt32Value)1u,
                    Name = (StringValue)((sheetName == null) ? "mySheet" : sheetName)
                };
                sheets!.Append(sheet);
                BuildExcelRowsWithHeader(sheetData, header, data);
                workbookPart.Workbook.Save();
            }
        }

        private void BuildExcelRowsWithHeader(SheetData sheetData, List<string> header, dynamic data)
        {
            Row row = new Row();
            List<string> list = new List<string>();
            foreach (string item in header)
            {
                list.Add(item);
                Cell cell = new Cell();
                cell.DataType = (EnumValue<CellValues>)CellValues.String;
                cell.CellValue = new CellValue(item);
                row.AppendChild(cell);
            }

            sheetData.AppendChild(row);
            foreach (object datum in data)
            {
                List<object> obj = (List<object>)(dynamic)datum;
                Row row2 = new Row();
                foreach (object item2 in obj)
                {
                    Cell cell2 = new Cell();
                    cell2.DataType = (EnumValue<CellValues>)CellValues.String;
                    if (item2 != null)
                    {
                        cell2.CellValue = new CellValue(item2.ToString());
                    }
                    else
                    {
                        cell2.CellValue = new CellValue(null);
                    }

                    row2.AppendChild(cell2);
                }

                sheetData.AppendChild(row2);
            }
        }

        private void BuildExcelRows(SheetData sheetData, DataTable table)
        {
            Row row = new Row();
            List<string> list = new List<string>();
            foreach (DataColumn column in table.Columns)
            {
                list.Add(column.ColumnName);
                Cell cell = new Cell();
                cell.DataType = (EnumValue<CellValues>)CellValues.String;
                cell.CellValue = new CellValue(column.ColumnName);
                row.AppendChild(cell);
            }

            sheetData.AppendChild(row);
            foreach (DataRow row3 in table.Rows)
            {
                Row row2 = new Row();
                foreach (string item in list)
                {
                    Cell cell2 = new Cell();
                    cell2.DataType = (EnumValue<CellValues>)CellValues.String;
                    cell2.CellValue = new CellValue(row3[item].ToString());
                    row2.AppendChild(cell2);
                }

                sheetData.AppendChild(row2);
            }
        }

        public void ToExcel1(DataTable table, string fileName)
        {
            DateTime.Now.ToString().Replace("/", "_").Replace(":", "_");
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            File.Create(fileName);
            using SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
            CreatePartsForExcel(document, table);
        }

        private void CreatePartsForExcel(SpreadsheetDocument document, DataTable data)
        {
            SheetData sheetData = GenerateSheetdataForDetails(data);
            WorkbookPart workbookPart = document.AddWorkbookPart();
            GenerateWorkbookPartContent(workbookPart);
            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPartContent(workbookStylesPart);
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPartContent(worksheetPart, sheetData);
        }

        private void GenerateWorksheetPartContent(WorksheetPart worksheetPart1, SheetData sheetData1)
        {
            Worksheet worksheet = new Worksheet
            {
                MCAttributes = new MarkupCompatibilityAttributes
                {
                    Ignorable = (StringValue)"x14ac"
                }
            };
            worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension = new SheetDimension
            {
                Reference = (StringValue)"A1"
            };
            SheetViews sheetViews = new SheetViews();
            SheetView sheetView = new SheetView
            {
                TabSelected = (BooleanValue)true,
                WorkbookViewId = (UInt32Value)0u
            };
            Selection selection = new Selection
            {
                ActiveCell = (StringValue)"A1",
                SequenceOfReferences = new ListValue<StringValue>
                {
                    InnerText = "A1"
                }
            };
            sheetView.Append(selection);
            sheetViews.Append(sheetView);
            SheetFormatProperties sheetFormatProperties = new SheetFormatProperties
            {
                DefaultRowHeight = (DoubleValue)15.0,
                DyDescent = (DoubleValue)0.25
            };
            PageMargins pageMargins = new PageMargins
            {
                Left = (DoubleValue)0.7,
                Right = (DoubleValue)0.7,
                Top = (DoubleValue)0.75,
                Bottom = (DoubleValue)0.75,
                Header = (DoubleValue)0.3,
                Footer = (DoubleValue)0.3
            };
            worksheet.Append(sheetDimension);
            worksheet.Append(sheetViews);
            worksheet.Append(sheetFormatProperties);
            worksheet.Append(sheetData1);
            worksheet.Append(pageMargins);
            worksheetPart1.Worksheet = worksheet;
        }

        private void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet = new Stylesheet
            {
                MCAttributes = new MarkupCompatibilityAttributes
                {
                    Ignorable = (StringValue)"x14ac"
                }
            };
            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            Fonts fonts = new Fonts
            {
                Count = (UInt32Value)2u,
                KnownFonts = (BooleanValue)true
            };
            Font font = new Font();
            FontSize fontSize = new FontSize
            {
                Val = (DoubleValue)11.0
            };
            Color color = new Color
            {
                Theme = (UInt32Value)1u
            };
            FontName fontName = new FontName
            {
                Val = (StringValue)"Calibri"
            };
            FontFamilyNumbering fontFamilyNumbering = new FontFamilyNumbering
            {
                Val = (Int32Value)2
            };
            FontScheme fontScheme = new FontScheme
            {
                Val = (EnumValue<FontSchemeValues>)FontSchemeValues.Minor
            };
            font.Append(fontSize);
            font.Append(color);
            font.Append(fontName);
            font.Append(fontFamilyNumbering);
            font.Append(fontScheme);
            Font font2 = new Font();
            Bold bold = new Bold();
            FontSize fontSize2 = new FontSize
            {
                Val = (DoubleValue)11.0
            };
            Color color2 = new Color
            {
                Theme = (UInt32Value)1u
            };
            FontName fontName2 = new FontName
            {
                Val = (StringValue)"Calibri"
            };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering
            {
                Val = (Int32Value)2
            };
            FontScheme fontScheme2 = new FontScheme
            {
                Val = (EnumValue<FontSchemeValues>)FontSchemeValues.Minor
            };
            font2.Append(bold);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);
            fonts.Append(font);
            fonts.Append(font2);
            Fills fills = new Fills
            {
                Count = (UInt32Value)2u
            };
            Fill fill = new Fill();
            PatternFill patternFill = new PatternFill
            {
                PatternType = (EnumValue<PatternValues>)PatternValues.None
            };
            fill.Append(patternFill);
            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill
            {
                PatternType = (EnumValue<PatternValues>)PatternValues.Gray125
            };
            fill2.Append(patternFill2);
            fills.Append(fill);
            fills.Append(fill2);
            Borders borders = new Borders
            {
                Count = (UInt32Value)2u
            };
            Border border = new Border();
            LeftBorder leftBorder = new LeftBorder();
            RightBorder rightBorder = new RightBorder();
            TopBorder topBorder = new TopBorder();
            BottomBorder bottomBorder = new BottomBorder();
            DiagonalBorder diagonalBorder = new DiagonalBorder();
            border.Append(leftBorder);
            border.Append(rightBorder);
            border.Append(topBorder);
            border.Append(bottomBorder);
            border.Append(diagonalBorder);
            Border border2 = new Border();
            LeftBorder leftBorder2 = new LeftBorder
            {
                Style = (EnumValue<BorderStyleValues>)BorderStyleValues.Thin
            };
            Color color3 = new Color
            {
                Indexed = (UInt32Value)64u
            };
            leftBorder2.Append(color3);
            RightBorder rightBorder2 = new RightBorder
            {
                Style = (EnumValue<BorderStyleValues>)BorderStyleValues.Thin
            };
            Color color4 = new Color
            {
                Indexed = (UInt32Value)64u
            };
            rightBorder2.Append(color4);
            TopBorder topBorder2 = new TopBorder
            {
                Style = (EnumValue<BorderStyleValues>)BorderStyleValues.Thin
            };
            Color color5 = new Color
            {
                Indexed = (UInt32Value)64u
            };
            topBorder2.Append(color5);
            BottomBorder bottomBorder2 = new BottomBorder
            {
                Style = (EnumValue<BorderStyleValues>)BorderStyleValues.Thin
            };
            Color color6 = new Color
            {
                Indexed = (UInt32Value)64u
            };
            bottomBorder2.Append(color6);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();
            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);
            borders.Append(border);
            borders.Append(border2);
            CellStyleFormats cellStyleFormats = new CellStyleFormats
            {
                Count = (UInt32Value)1u
            };
            CellFormat cellFormat = new CellFormat
            {
                NumberFormatId = (UInt32Value)0u,
                FontId = (UInt32Value)0u,
                FillId = (UInt32Value)0u,
                BorderId = (UInt32Value)0u
            };
            cellStyleFormats.Append(cellFormat);
            CellFormats cellFormats = new CellFormats
            {
                Count = (UInt32Value)3u
            };
            CellFormat cellFormat2 = new CellFormat
            {
                NumberFormatId = (UInt32Value)0u,
                FontId = (UInt32Value)0u,
                FillId = (UInt32Value)0u,
                BorderId = (UInt32Value)0u,
                FormatId = (UInt32Value)0u
            };
            CellFormat cellFormat3 = new CellFormat
            {
                NumberFormatId = (UInt32Value)0u,
                FontId = (UInt32Value)0u,
                FillId = (UInt32Value)0u,
                BorderId = (UInt32Value)1u,
                FormatId = (UInt32Value)0u,
                ApplyBorder = (BooleanValue)true
            };
            CellFormat cellFormat4 = new CellFormat
            {
                NumberFormatId = (UInt32Value)0u,
                FontId = (UInt32Value)1u,
                FillId = (UInt32Value)0u,
                BorderId = (UInt32Value)1u,
                FormatId = (UInt32Value)0u,
                ApplyFont = (BooleanValue)true,
                ApplyBorder = (BooleanValue)true
            };
            cellFormats.Append(cellFormat2);
            cellFormats.Append(cellFormat3);
            cellFormats.Append(cellFormat4);
            CellStyles cellStyles = new CellStyles
            {
                Count = (UInt32Value)1u
            };
            CellStyle cellStyle = new CellStyle
            {
                Name = (StringValue)"Normal",
                FormatId = (UInt32Value)0u,
                BuiltinId = (UInt32Value)0u
            };
            cellStyles.Append(cellStyle);
            DifferentialFormats differentialFormats = new DifferentialFormats
            {
                Count = (UInt32Value)0u
            };
            TableStyles tableStyles = new TableStyles
            {
                Count = (UInt32Value)0u,
                DefaultTableStyle = (StringValue)"TableStyleMedium2",
                DefaultPivotStyle = (StringValue)"PivotStyleLight16"
            };
            StylesheetExtensionList stylesheetExtensionList = new StylesheetExtensionList();
            StylesheetExtension stylesheetExtension = new StylesheetExtension
            {
                Uri = (StringValue)"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"
            };
            stylesheetExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            StylesheetExtension stylesheetExtension2 = new StylesheetExtension
            {
                Uri = (StringValue)"{9260A510-F301-46a8-8635-F512D64BE5F5}"
            };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            stylesheetExtensionList.Append(stylesheetExtension);
            stylesheetExtensionList.Append(stylesheetExtension2);
            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(cellFormats);
            stylesheet.Append(cellStyles);
            stylesheet.Append(differentialFormats);
            stylesheet.Append(tableStyles);
            stylesheet.Append(stylesheetExtensionList);
            workbookStylesPart1.Stylesheet = stylesheet;
        }

        private void GenerateWorkbookPartContent(WorkbookPart workbookPart1)
        {
            Workbook workbook = new Workbook();
            Sheets sheets = new Sheets();
            Sheet sheet = new Sheet
            {
                Name = (StringValue)"Sheet1",
                SheetId = (UInt32Value)1u,
                Id = (StringValue)"rId1"
            };
            sheets.Append(sheet);
            workbook.Append(sheets);
            workbookPart1.Workbook = workbook;
        }

        private SheetData GenerateSheetdataForDetails(DataTable data)
        {
            return new SheetData();
        }
    }
}
