using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelStringSearch
{
    /// <summary>
    /// 共有文字列パーツ(SharedStringsPart）を扱うクラス
    /// </summary>
    internal class OSNSharedStrings
    {
        private readonly SharedStringTablePart SharedStringTablePart;
        private readonly SharedStringTable SharedStringTable;

        internal Dictionary<uint, SharedStringItem> IndexSiTable { get; private set; }
        internal Dictionary<string, List<uint>> SiTextIndexListTable { get; private set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="sharedStringTablePart">Excelファイルの共有文字列パーツ</param>
        internal OSNSharedStrings(SharedStringTablePart sharedStringTablePart)
        {
            this.IndexSiTable = new Dictionary<uint, SharedStringItem>();
            this.SiTextIndexListTable = new Dictionary<string, List<uint>>();
            this.SharedStringTablePart = sharedStringTablePart;
            this.SharedStringTable = sharedStringTablePart?.SharedStringTable;
            this.ParseIndexSiTable();
        }

        /// <summary>
        /// si要素をパースして、indexと文字列の連想配列を作る
        /// </summary>
        private void ParseIndexSiTable()
        {
            if (this.SharedStringTable == null) return;
            this.IndexSiTable.Clear();
            this.SiTextIndexListTable.Clear();
            uint siIndex = 0;
            foreach (var si in this.SharedStringTable.OfType<SharedStringItem>())
            {
                this.IndexSiTable.Add(siIndex, si);
                if (!this.SiTextIndexListTable.ContainsKey(si.InnerText))
                {
                    this.SiTextIndexListTable.Add(si.InnerText, new List<uint>());
                }
                this.SiTextIndexListTable[si.InnerText].Add(siIndex++);
            }
        }

        /// <summary>
        /// 文字列のindexの集合を取得する
        /// </summary>
        /// <param name="str">検索対象の文字列</param>
        /// <returns>検索対象の文字列が含まれるsi要素のindexの集合</returns>
        internal List<uint> GetStringIndexSet(string str)
        {
            var ret = new List<uint>();
            foreach (var siTextIndexList in this.SiTextIndexListTable)
            {
                var siText = siTextIndexList.Key;
                if (!siText.Contains(str)) continue;

                ret.AddRange(this.SiTextIndexListTable[siText]);
            }

            return ret;
        }
    }
}
