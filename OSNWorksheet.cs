using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelStringSearch
{
    class OSNWorksheet
    {
        public Sheet Sheet { get; private set; }

        private WorksheetPart WorksheetPart;

        public HashSet<Cell> CellSet { get; private set; }

        public OSNWorksheet(Sheet sheet, WorksheetPart worksheetPart)
        {
            this.Sheet = sheet;
            this.WorksheetPart = worksheetPart;
            this.CellSet = new HashSet<Cell>();
            this.ParseCell();
        }

        private void ParseCell()
        {
            if (this.CellSet.Any()) this.CellSet.Clear();
            var worksheet = this.WorksheetPart.Worksheet;
            foreach (var cell in this.WorksheetPart.Worksheet.Descendants<Cell>())
            {
                this.CellSet.Add(cell);
            }
        }

        public Dictionary<int, HashSet<Cell>> GetStrIndexCellSetTable()
        {
            var strIndexCellSetTable = new Dictionary<int, HashSet<Cell>>();
            if (!this.CellSet.Any()) this.ParseCell();
            foreach (var cell in this.CellSet)
            {
                if (cell.DataType != CellValues.SharedString) continue;
                int index = System.Int32.Parse(cell.InnerText);
                if (strIndexCellSetTable.ContainsKey(index))
                {
                    strIndexCellSetTable[index].Add(cell);
                }
                else
                {
                    strIndexCellSetTable.Add(index, new HashSet<Cell>() { cell });
                }
            }
            return strIndexCellSetTable;
        }
    }
}
