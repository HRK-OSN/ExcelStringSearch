using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelStringSearch
{
    public class SearchString
    {
        private readonly OSNWorkbook OSNWorkbook;

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
                var sharedStringIndexCellSetTable = osnWorksheet.SharedStringIndexCellSetTable;
                var siIndexCellSetTable = new Dictionary<uint, HashSet<Cell>>(sharedStringIndexCellSetTable.Count);
                foreach (var targetSiIndex in targetSiIndexSet)
                {
                    if (sharedStringIndexCellSetTable.TryGetValue(targetSiIndex, out var cellSet))
                    {
                        if (!siIndexCellSetTable.ContainsKey(targetSiIndex))
                        {
                            siIndexCellSetTable.Add(targetSiIndex, new HashSet<Cell>());
                        }
                        foreach (var cell in cellSet)
                        {
                            siIndexCellSetTable[targetSiIndex].Add(cell);
                        }
                    }
                }

                if (!siIndexCellSetTable.Any()) continue;

                var searchStringResult = new SearchStringResult
                {
                    OSNWorksheet = osnWorksheet,
                    SiIndexCellSetTable = siIndexCellSetTable
                };
                result.Add(searchStringResult);
            }
            return result;
        }

        private HashSet<uint> SearchSSIIndex(string targetString)
        {
            var indexSiTable = this.OSNWorkbook.OSNSharedStrings.IndexSiTable;
            var stringIndex = new HashSet<uint>(indexSiTable.Count);

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
