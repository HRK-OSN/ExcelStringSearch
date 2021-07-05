using System.Collections.Generic;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelStringSearch
{
    internal class OSNWorkbook
    {
        private readonly MemoryStream WorkbookStream;
        private readonly SpreadsheetDocument XlsxDocument;
        private WorkbookPart WorkbookPart;
        private Workbook Workbook;

        private Sheets Sheets;
        internal List<OSNWorksheet> OSNWorksheetList { get; private set; }

        internal OSNSharedStrings OSNSharedStrings { get; private set; }

        internal OSNWorkbook(Stream stream)
        {
            byte[] buffer = new byte[stream.Length];
            stream.Read(buffer, 0, buffer.Length);
            this.WorkbookStream = new MemoryStream();
            this.WorkbookStream.Write(buffer, 0, buffer.Length);

            stream.Position = 0;

            this.XlsxDocument = SpreadsheetDocument.Open(this.WorkbookStream, true);
            this.Init();
        }

        private void Init()
        {
            this.WorkbookPart = this.XlsxDocument.WorkbookPart;
            this.Workbook = this.WorkbookPart.Workbook;
            this.Sheets = this.Workbook.Sheets;

            this.InitWorkSheets();
            this.InitRelatedParts();
        }

        private void InitWorkSheets()
        {
            this.Sheets = this.Workbook.Sheets;

            var sheetList = this.Sheets.OfType<Sheet>();
            this.OSNWorksheetList = new List<OSNWorksheet>(sheetList.Count());
            foreach (var sheet in sheetList)
            {
                this.OSNWorksheetList.Add(new OSNWorksheet(sheet, (WorksheetPart)this.WorkbookPart.GetPartById(sheet.Id)));
            }
        }

        private void InitRelatedParts()
        {
            this.OSNSharedStrings = new OSNSharedStrings(this.WorkbookPart.SharedStringTablePart);
        }
    }
}
