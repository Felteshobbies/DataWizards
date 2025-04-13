using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace libDataWizard
{
    public class XLS
    {
        public Boolean overwrite { get; set; }
        public string filePath { get; set; }

        public XLS(string filePath, bool overwrite)
        {
            this.filePath = filePath;
            this.overwrite = overwrite;

            if (File.Exists(filePath) && !overwrite)
            {
                throw new Exception("File exists.");               
            } 

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                // Workbook und Worksheet erstellen
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Tabelle1"
                };
                sheets.Append(sheet);

                // Tabellendaten hinzufügen
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Beispielhafte Zeilen und Zellen
                Row row1 = new Row();
                row1.Append(CreateTextCell("Text in Zelle A1"));
                row1.Append(CreateTextCell("Text in Zelle B1"));

                sheetData.Append(row1);

                Row row2 = new Row();
                row2.Append(CreateTextCell("Noch ein Text"));
                row2.Append(CreateTextCell("Mehr Text"));
                sheetData.Append(row2);

                workbookPart.Workbook.Save();
            }
        }


        private Cell CreateTextCell(string cellValue)
        {



            return new Cell()
            {
                //CellReference = cellReference,
                DataType = CellValues.String,
                CellValue = new CellValue(cellValue)
            };
        }
    }




  


}
