using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelStringSearch
{
    public class OSNSharedStrings
    {
        private readonly SharedStringTablePart SharedStringTablePart;
        private readonly SharedStringTable SharedStringTable;

        public Dictionary<uint, SharedStringItem> IndexSiTable { get; private set; }
        public Dictionary<string, uint> SiTextIndexTable { get; private set; }

        public OSNSharedStrings(SharedStringTablePart sharedStringTablePart)
        {
            this.IndexSiTable = new Dictionary<uint, SharedStringItem>();
            this.SiTextIndexTable = new Dictionary<string, uint>();
            if (sharedStringTablePart == null) return;
            this.SharedStringTablePart = sharedStringTablePart;
            this.SharedStringTable = sharedStringTablePart.SharedStringTable;
            this.ParseIndexSiTable();
        }

        private void ParseIndexSiTable()
        {
            uint siIndex = 0;
            using var reader = OpenXmlReader.Create(this.SharedStringTable);
            while (reader.Read())
            {
                if (reader.ElementType == typeof(SharedStringTable)) continue;
                do
                {
                    if (reader.ElementType == typeof(SharedStringItem))
                    {
                        var si = (SharedStringItem)reader.LoadCurrentElement();
                        this.IndexSiTable.Add(siIndex, si);
                        this.SiTextIndexTable.Add(si.InnerText, siIndex++);
                    }
                } while (reader.ReadNextSibling());
            }
        }
    }
}
