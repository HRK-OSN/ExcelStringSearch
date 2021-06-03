using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelStringSearch
{
    public class SearchStringResult
    {
        public OSNWorksheet OSNWorksheet { get; set; }

        public Dictionary<uint, HashSet<Cell>> SiIndexCellSetTable { get; set; }

        public SearchStringResult()
        {
            this.SiIndexCellSetTable = new Dictionary<uint, HashSet<Cell>>();
        }
    }
}
