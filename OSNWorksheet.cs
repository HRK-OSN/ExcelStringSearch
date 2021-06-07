using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelStringSearch
{
    public class OSNWorksheet
    {
        private readonly Sheet Sheet;

        private readonly WorksheetPart WorksheetPart;
        private readonly Worksheet Worksheet;

        public Dictionary<uint, HashSet<Cell>> SharedStringIndexCellSetTable { get; private set; }

        public string Name
        {
            get { return this.Sheet.Name; }
            set { this.Sheet.Name = value; }
        }

        public OSNWorksheet(Sheet sheet, WorksheetPart worksheetPart)
        {
            this.Sheet = sheet;
            this.WorksheetPart = worksheetPart;
            this.Worksheet = worksheetPart.Worksheet;
            this.InitSharedStringIndexCell();
        }

        private void InitSharedStringIndexCell()
        {
            this.SharedStringIndexCellSetTable = new Dictionary<uint, HashSet<Cell>>();
            using var reader = OpenXmlReader.Create(this.Worksheet);
            while (reader.Read())
            {
                if (reader.ElementType != typeof(Cell)) continue;
                var cell = (Cell)reader.LoadCurrentElement();
                if (cell.DataType == null || cell.DataType != CellValues.SharedString) continue;
                uint index = System.UInt32.Parse(cell.InnerText);
                if (!this.SharedStringIndexCellSetTable.ContainsKey(index))
                {
                    this.SharedStringIndexCellSetTable.Add(index, new HashSet<Cell>());
                }
                this.SharedStringIndexCellSetTable[index].Add(cell);
            }
        }

    }
}
