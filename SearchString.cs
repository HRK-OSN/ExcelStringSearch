using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelStringSearch
{
    public class SearchString
    {
        private OSNSharedStrings OSNSharedStrings;
        private OSNWorkbook OSNWorkbook;

        public SearchString(Stream stream)
        {
            this.OSNWorkbook = new OSNWorkbook(stream);
        }

        public List<SearchStringResult> Search(string targetString)
        {
            var targetSiIndexSet = this.SearchSSIIndex(targetString);
            var result = new List<SearchStringResult>(this.OSNWorkbook.OSNWorksheetTable.Count);

            foreach (var osnWorksheet in this.OSNWorkbook.OSNWorksheetTable.Values)
            {
                var strIndexCellSetTable = osnWorksheet.GetStrIndexCellSetTable();
                var siIndexCellSetTable = new Dictionary<int, HashSet<Cell>>(strIndexCellSetTable.Count);
                foreach (var targetSiIndex in targetSiIndexSet)
                {
                    if (!strIndexCellSetTable.ContainsKey(targetSiIndex)) continue;
                    siIndexCellSetTable.Add(targetSiIndex, strIndexCellSetTable[targetSiIndex]);
                }

                if (!siIndexCellSetTable.Any()) continue;

                var searchStringResult = new SearchStringResult();
                searchStringResult.OSNWorksheet = osnWorksheet;
                searchStringResult.SiIndexCellSetTable = siIndexCellSetTable;
                result.Add(searchStringResult);
            }

            return result;
        }

        private HashSet<int> SearchSSIIndex(string targetString)
        {
            var indexSiTable = this.OSNWorkbook.OSNSharedStrings.IndexSiTable;
            var stringIndex = new HashSet<int>(indexSiTable.Count);

            foreach (var indexSi in indexSiTable)
            {
                var ssiText = indexSi.Value.InnerText;
                if (!ssiText.Contains(targetString)) continue;
                stringIndex.Add(indexSi.Key);
            }

            return stringIndex;
        }
    }
}
