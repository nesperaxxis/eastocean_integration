using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    class ItemCount : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "stockcount";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_STOCK_COUNT;
        internal const int _sapObjType = (int)SAPbobsCOM.BoObjectTypes.oStockTakings;
        internal const string _keyField = "DocEntry";
        internal const string _nameField = "DocNum";
        internal const string _filterField = _nameField;


        internal override string wmsAPIEndPoint => _wmsAPIEndPoint;
        internal override string wmsObjectType => _objectType;
        internal override int sapObjectType => _sapObjType;
        internal override string sapKeyField => _keyField;
        internal override string sapNameField => _nameField;
        int _docEntry = 0;

        public string ObjType = _objectType;
        public int DocumentKey;
        public int DocumentNo;
        public string CountDate;
        public string Status;
        public string Ref2 = "";
        public string Remarks ="";
       public string DocType = "ITEM_COUNTING";

        public List<INCLine> Lines = new List<INCLine>();

        public ItemCount(int docEntry) : base()
        {
            _docEntry = docEntry;

            var sql = String.Format(Src.Resource.Queries.GET_RECORD_OINC, _docEntry);
            var rs = eCommon.ExecuteQuery(sql);
            DocumentKey = (int)rs.Fields.Item("DocEntry").Value;
            DocumentNo = (int)rs.Fields.Item("DocNum").Value;
            CountDate = ((DateTime)rs.Fields.Item("CountDate").Value).ToString("yyyyMMdd");
            Status = rs.Fields.Item("Status").Value.ToString();
            Ref2 = String.IsNullOrEmpty(rs.Fields.Item("Ref2").Value.ToString()) ? "" : rs.Fields.Item("Ref2").Value.ToString();
            Remarks = String.IsNullOrEmpty(rs.Fields.Item("Remarks").Value.ToString()) ? "" : rs.Fields.Item("Remarks").Value.ToString();
            //CardCode = rs.Fields.Item("CardCode").Value.ToString();

            Lines = INCLine.GetItems(_docEntry);
        }

        internal override string sapKeyVal => this.DocumentKey.ToString();
        internal override string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }
    }

    public class INCLine
    {
        private static int objType = (int)SAPbobsCOM.BoObjectTypes.oStockTakings;
        public int DocEntry;
        public int LineNum;
        public string ItemCode;
        public string ItemName;
        public string UOM;
        public double InWhsQty;
        public string IsCounted;
        public double CountQty;
        public string WhsCode = "";
        public string BinCode = "";
        public string SnBCode = "";

        public INCLine() { }

        public static List<INCLine> GetItems(int docEntry)
        {
            List<INCLine> lines = new List<INCLine>();
            string tableName = eCommon.GetTableName(objType.ToString());
            string sql = String.Format(Resource.Queries.GET_RECORD_INC1, docEntry);
            var rs = eCommon.ExecuteQuery(sql);
            for (int i = 0; i < rs.RecordCount; i++)
            {
                lines.Add(new INCLine
                {
                    DocEntry = (int)rs.Fields.Item("DocEntry").Value,
                    LineNum = (int)rs.Fields.Item("LineNum").Value,
                    ItemCode = rs.Fields.Item("ItemCode").Value.ToString(),
                    ItemName = rs.Fields.Item("ItemName").Value.ToString(),
                    InWhsQty = (double)rs.Fields.Item("InWhsQty").Value,
                    IsCounted = rs.Fields.Item("IsCounted").Value.ToString(),
                    CountQty = (double)rs.Fields.Item("CountQty").Value,
                    UOM = rs.Fields.Item("InvntryUom").Value.ToString(),
                    WhsCode = rs.Fields.Item("WhsCode").Value.ToString(),             
                    BinCode = rs.Fields.Item("BinCode").Value.ToString(),
                    SnBCode = rs.Fields.Item("SnBCode").Value.ToString()
                }
                );
                rs.MoveNext();
            }

            return lines;
        }
    }
}
