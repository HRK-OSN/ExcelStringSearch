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
        private MemoryStream WorkbookStream;
        private SpreadsheetDocument XlsxDocument;
        private WorkbookPart WorkbookPart;

        public OSNSharedStrings OSNSharedStrings { get; private set; }
        public Dictionary<int, OSNWorksheet> OSNWorksheetTable { get; private set; }

        public OSNWorkbook(Stream stream)
        {
            byte[] buffer = new byte[stream.Length];
            stream.Read(buffer, 0, buffer.Length);
            this.WorkbookStream = new MemoryStream();
            this.WorkbookStream.Write(buffer, 0, buffer.Length);

            stream.Position = 0;

            this.XlsxDocument = SpreadsheetDocument.Open(this.WorkbookStream, true);

            this.OSNWorksheetTable = new Dictionary<int, OSNWorksheet>();
        }

        public void ParseDocument()
        {
            this.ParseWorkbook();
            this.ParseRelatedParts();
        }

        private void ParseWorkbook()
        {
            this.WorkbookPart = this.XlsxDocument.WorkbookPart;
        }

        private void ParseRelatedParts()
        {
            var relatedParts = this.WorkbookPart.Parts;
            foreach (var relatedPart in relatedParts)
            {
                switch (relatedPart.OpenXmlPart)
                {
                    case SharedStringTablePart sharedStringTablePart:
                        this.OSNSharedStrings = new OSNSharedStrings(sharedStringTablePart);
                        break;
                }
            }
        }

        private void ParseOSNWorksheetTable()
        {
            var workbook = this.WorkbookPart.Workbook;
            if (this.OSNWorksheetTable.Any()) this.OSNWorksheetTable.Clear();
            int localSheetIndex = 0;
            using var reader = OpenXmlReader.Create(workbook.Sheets);
            while (reader.Read())
            {
                var sheet = (Sheet)reader.LoadCurrentElement();
                this.OSNWorksheetTable.Add(localSheetIndex++,
                    new OSNWorksheet(sheet, (WorksheetPart)this.WorkbookPart.GetPartById(sheet.Id)));
            }
        }
    }
}
