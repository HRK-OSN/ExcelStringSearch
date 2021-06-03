using System.Collections.Generic;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelStringSearch
{
    public class OSNWorkbook
    {
        private readonly MemoryStream WorkbookStream;
        private readonly SpreadsheetDocument XlsxDocument;
        private WorkbookPart WorkbookPart;
        private Workbook Workbook;

        private Sheets Sheets;
        public Dictionary<uint, Sheet> LocalIndexSheetTable { get; private set; }
        public Dictionary<Sheet, OSNWorksheet> OSNWorksheetTable { get; private set; }

        public OSNSharedStrings OSNSharedStrings { get; private set; }

        public OSNWorkbook(Stream stream)
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
            this.LocalIndexSheetTable = new Dictionary<uint, Sheet>();
            this.OSNWorksheetTable = new Dictionary<Sheet, OSNWorksheet>();

            using var reader = OpenXmlReader.Create(this.Sheets);
            uint localIndex = 0;
            while (reader.Read())
            {
                if (reader.ElementType == typeof(Sheets)) continue;
                do
                {
                    if (reader.ElementType == typeof(Sheet))
                    {
                        var sheet = (Sheet)reader.LoadCurrentElement();
                        this.LocalIndexSheetTable.Add(localIndex++, sheet);
                        this.OSNWorksheetTable.Add(sheet, new OSNWorksheet(sheet, (WorksheetPart)this.WorkbookPart.GetPartById(sheet.Id)));
                    }
                } while (reader.ReadNextSibling());
            }
        }

        private void InitRelatedParts()
        {
            this.OSNSharedStrings = new OSNSharedStrings(this.WorkbookPart.SharedStringTablePart);
        }
    }
}
