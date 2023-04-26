using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Dynamic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Branch.Tools
{
    internal class OpenXmlHandler
    {
        public SpreadsheetDocument spreadsheetDocument { get; }
        private WorkbookPart workbookPart;
        private Workbook workbook;
        //private Sheets sheets;
        //private SharedStringTable sharedStringTable;
        public OpenXmlHandler(string path)
        {
            spreadsheetDocument = SpreadsheetDocument.Open(path, true);
            workbookPart = spreadsheetDocument.WorkbookPart;
            workbook = workbookPart.Workbook;
        }
        public OpenXmlHandler(SpreadsheetDocument spreadsheetDocument)
        {
            this.spreadsheetDocument = spreadsheetDocument;
            workbookPart = spreadsheetDocument.WorkbookPart;
            workbook = workbookPart.Workbook;
        }
        /// <summary>
        /// 创建一个空白工作簿，包含一个空白工作表
        /// </summary>
        /// <param name="path"></param>
        /// <param name="sheetId"></param>
        /// <param name="SheetName"></param>
        public static SpreadsheetDocument CreateWorkbook(string path)
        {
            //创建工作簿文件
            //默认 AutoSave = true ; Editable = true ; Type = xlsx
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);

            //添加 workbookpart
            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            //添加 worksheetpart 到 workbookpart
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            ////添加 sheets 到 workbook
            //Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

            ////添加 worksheet 到 workbook 并关联
            //Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = sheetId, Name = SheetName };
            //sheets.Append(sheet);

            //workbookPart.Workbook.Save();
            return spreadsheetDocument;
        }

        public List<List<string>> ReadSheet(uint index = 1)
        {
            IEnumerable<WorksheetPart> worksheetParts = workbookPart.WorksheetParts;
            UInt32Value sheetIndex = UInt32Value.FromUInt32(index);

            Sheets sheets = workbook.Sheets;
            if (sheets == null) return null;

            SharedStringTable sharedStringTable = null;
            try
            {
                sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;
            }
            catch (System.NullReferenceException) { }

            Sheet sheet = sheets.Elements<Sheet>().FirstOrDefault(x => x.SheetId == sheetIndex);
            if (sheet == null) return null;

            string relationshipId = sheet.Id.ToString();

            WorksheetPart worksheetPart = worksheetParts.FirstOrDefault(x => workbookPart.GetIdOfPart(x) == relationshipId);
            if (worksheetPart == null) return null;

            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            List<List<string>> result = new List<List<string>>();

            foreach (Row row in sheetData.Elements<Row>())
            {
                List<string> strings = new List<string>();
                foreach (Cell cell in row.Elements<Cell>())
                {
                    string text = cell.CellValue.Text;
                    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                    {
                        var xmlPart = sharedStringTable.ElementAt(Convert.ToInt32(text));//NullReferenceException??
                        text = xmlPart.FirstChild.InnerText;
                    }
                    strings.Add(text);
                }
                result.Add(strings);
            }
            return result;
        }
        public List<List<string>> ReadSheet(string name)
        {
            IEnumerable<WorksheetPart> worksheetParts = workbookPart.WorksheetParts;
            Sheets sheets = workbook.Sheets;
            if (sheets == null) return null;

            Sheet sheet = sheets.Elements<Sheet>().FirstOrDefault(x => x.Name == name);
            if (sheet == null) return null;

            SharedStringTable sharedStringTable = null;
            try
            {
                sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;
            }
            catch (System.NullReferenceException) { }

            string relationshipId = sheet.Id.ToString();

            WorksheetPart worksheetPart = worksheetParts.FirstOrDefault(x => workbookPart.GetIdOfPart(x) == relationshipId);
            if (worksheetPart == null) return null;

            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            List<List<string>> result = new List<List<string>>();
            foreach (Row row in sheetData.Elements<Row>())
            {
                List<string> strings = new List<string>();
                foreach (Cell cell in row.Elements<Cell>())
                {
                    string text = cell.CellValue.Text;
                    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                    {
                        var xmlPart = sharedStringTable.ElementAt(Convert.ToInt32(text));
                        text = xmlPart.FirstChild.InnerText;
                    }
                    strings.Add(text);
                }
                result.Add(strings);
            }
            return result;
        }
        public void AddSheet(List<List<string>> lists, string name = null)
        {
            if (lists == null || lists.Count == 0) return;

            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.Sheets;

            if (sheets == null)
            {
                sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
            }

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            SheetData sheetData = new SheetData();

            #region fil sheetData
            uint rowReference = 1;
            lists.ForEach(x =>
                {
                    Row row = new Row() { RowIndex = rowReference };
                    int cell_idnex = 0;
                    x.ForEach(y =>
                        {
                            string cellReference_prefix = IterationLetter(cell_idnex);
                            string cellReference = cellReference_prefix + rowReference;
                            Cell cell = new Cell(new CellValue(y)) { DataType = new EnumValue<CellValues>(CellValues.String), CellReference = cellReference };
                            row.Append(cell);
                            cell_idnex++;
                        });
                    sheetData.Append(row);
                    rowReference++;
                });
            #endregion

            Worksheet worksheet = new Worksheet(sheetData);
            worksheetPart.Worksheet = worksheet;
            uint index = 1;
            if (name == null)
            {
                name = "Sheet" + index.ToString();
                if (sheets.Count() > 0)
                {
                    while (sheets.Elements<Sheet>().Any(x => x.Name == name))
                    {
                        index++;
                        name = "Sheet" + index.ToString();
                    }
                }
            }
            uint sheetId = 1;
            if (sheets.Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(x => x.SheetId.Value).Max() + 1;
            }
            Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = sheetId, Name = name };
            //sheets.InsertAt(sheet, ((int)index));
            sheets.Append(sheet);
        }
        public void TestAdd()
        {
            // Open the document for editing.

            // Add a blank WorksheetPart.
            WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Give the new worksheet a name.
            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);

        }
        public void ReOrderSheet()
        {
            Sheets orderedSheets = new Sheets();
            uint index = 1;
            foreach (Sheet sheet in workbook.Sheets)
            {
                Sheet newSheet = new Sheet() { Name = sheet.Name, SheetId = index, Id = sheet.Id };
                orderedSheets.Append(newSheet);
            }
            workbook.Sheets = orderedSheets;

        }
        public void GetAllShheet()
        {
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            IEnumerable<WorksheetPart> worksheetParts = workbookPart.WorksheetParts;
            MessageBox.Show($"{worksheetParts.Count()}");
        }

        public void Dispose()
        {
            workbook.Save();
            spreadsheetDocument.Dispose();
        }
        private string IterationLetter(int index)
        {
            string target = "";
            if (index / 26 > 0)
            {
                target += IterationLetter(index / 26 - 1);
            }

            int ascii_A = (int)Convert.ToByte('A');
            int target_index = ascii_A + index % 26;
            target += Convert.ToChar(target_index).ToString();
            return target;
        }
    }
}
