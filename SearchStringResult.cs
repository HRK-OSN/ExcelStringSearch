using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelStringSearch
{
    /// <summary>
    /// 文字の検索結果を表すクラス
    /// </summary>
    public class SearchStringResult
    {
        /// <summary>
        /// 検索対象の文字列
        /// </summary>
        private readonly string SearchTargetStr;

        /// <summary>
        /// 検索対象の文字列を含むシート名とそのcellの集合の連想配列
        /// </summary>
        public Dictionary<string, HashSet<Cell>> SheetNameCells { get; private set; }
        
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="searchTargetStr">検索対象の文字列</param>
        public SearchStringResult(string searchTargetStr)
        {
            this.SearchTargetStr = searchTargetStr;
            this.SheetNameCells = new Dictionary<string, HashSet<Cell>>();
        }

        internal void AddSheetNameCells(string sheetName, HashSet<Cell> cells)
        {
            this.SheetNameCells.Add(sheetName, cells);
        }
    }
}
