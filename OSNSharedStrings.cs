using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelStringSearch
{
    class OSNSharedStrings
    {
        private SharedStringTablePart SharedStringTablePart;

        public Dictionary<int, SharedStringItem> IndexSiTable { get; private set; }

        public bool HasPart()
        {
            return this.SharedStringTablePart != null ? true : false;
        }

        public OSNSharedStrings(SharedStringTablePart sharedStringTablePart)
        {
            this.SharedStringTablePart = sharedStringTablePart;
            this.IndexSiTable = new Dictionary<int, SharedStringItem>();
            this.ParseIndexSiTable();
        }

        private void ParseIndexSiTable()
        {
            if (!this.HasPart()) return;
            var sharedStringTable = this.SharedStringTablePart.SharedStringTable;
            int siIndex = 0;
            using var reader = OpenXmlReader.Create(sharedStringTable);
            while (reader.Read())
            {
                if (reader.ElementType != typeof(SharedStringItem)) continue;
                this.IndexSiTable.Add(siIndex++, (SharedStringItem)reader.LoadCurrentElement());
            }
        }

        public string GetIndexString(int index)
        {
            return this.IndexSiTable[index].InnerText;
        }
    }
}
