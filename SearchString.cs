using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelStringSearch
{
    /// <summary>
    /// Excelファイル内の文字列を検索する
    /// </summary>
    public class SearchString
    {
        /// <summary>
        /// 検索するExcelファイル
        /// </summary>
        private readonly OSNWorkbook OSNWorkbook;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="stream">検索対象のStream</param>
        public SearchString(Stream stream)
        {
            this.OSNWorkbook = new OSNWorkbook(stream);
        }

        /// <summary>
        /// 文字列を含むsheetとセルを取得
        /// </summary>
        /// <param name="targetString">検索対象の文字列</param>
        /// <returns>文字の検索結果を表すクラス</returns>
        public SearchStringResult Search(string targetString)
        {
            var targetSiIndexSet = this.OSNWorkbook.OSNSharedStrings.GetStringIndexSet(targetString);
            var result = new SearchStringResult(targetString);

            foreach (var osnWorksheet in this.OSNWorkbook.OSNWorksheetList)
            {
                var sharedStringIndexCellSetTable = osnWorksheet.SharedStringIndexCellSetTable;

                var cellList = new List<Cell>();
                foreach (var targetSiIndex in targetSiIndexSet)
                {
                    if (sharedStringIndexCellSetTable.TryGetValue(targetSiIndex, out var cells))
                    {
                        cellList.AddRange(cells);
                    }
                }
                if (!cellList.Any()) continue;
                result.AddSheetNameCells(osnWorksheet.Name, cellList.ToHashSet()s);
            }
            return result;
        }
    }
}
