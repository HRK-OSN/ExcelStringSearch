using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelStringSearch
{
    public class SearchStringResult
    {
        public OSNWorksheet OSNWorksheet { get; set; }

        public Dictionary<int, HashSet<Cell>> SiIndexCellSetTable { get; set; }

        public SearchStringResult()
        {
            this.SiIndexCellSetTable = new Dictionary<int, HashSet<Cell>>();
        }
    }
}
