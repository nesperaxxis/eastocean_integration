using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    class ARReturns : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "DoReturn";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_AR_RETURNS;
        internal const int _sapObjType = (int)SAPbobsCOM.BoObjectTypes.oReturns;
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
        public string PostDate;
        public string DueDate;
        public string CardCode;
        public string Location = "";
        public string DocType = "RETURNS_W_DO";

        public List<RDNLine> Lines = new List<RDNLine>();

        public ARReturns(int docEntry) : base()
        {
            _docEntry = docEntry;

            var sql = String.Format(Src.Resource.Queries.GET_RECORD_ORDN, _docEntry);
            var rs = eCommon.ExecuteQuery(sql);
            DocumentKey = (int)rs.Fields.Item("DocEntry").Value;
            DocumentNo = (int)rs.Fields.Item("DocNum").Value;
            DueDate = ((DateTime)rs.Fields.Item("DocDueDate").Value).ToString("yyyyMMdd");
            PostDate = ((DateTime)rs.Fields.Item("DocDate").Value).ToString("yyyyMMdd");
            CardCode = rs.Fields.Item("CardCode").Value.ToString();

            Lines = RDNLine.GetItems(_docEntry);
        }

        internal override string sapKeyVal => this.DocumentKey.ToString();
        internal override string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }
    }

    public class RDNLine
    {
        private static int objType = (int)SAPbobsCOM.BoObjectTypes.oReturns;
        public int DocEntry;
        public int LineNum;
        public string ItemCode;
        public string UOM;
        public double Quantity;
        public double InvQuantity;
        public string ItemName;
        public string Whse = "";
        public string SnBCode = "";
        public string BinCode = "";
        public string ReturnReason = "";

        public RDNLine() { }

        public static List<RDNLine> GetItems(int docEntry)
        {
            List<RDNLine> lines = new List<RDNLine>();
            string tableName = eCommon.GetTableName(objType.ToString());
            string sql = String.Format(Resource.Queries.OINM_GET_LINE_SNB_BIN_INFO, objType, docEntry, tableName);
            var rs = eCommon.ExecuteQuery(sql);
            for (int i = 0; i < rs.RecordCount; i++)
            {
                lines.Add(new RDNLine
                {
                    DocEntry = (int)rs.Fields.Item("DocEntry").Value,
                    LineNum = (int)rs.Fields.Item("LineKey").Value,
                    ItemCode = rs.Fields.Item("ItemCode").Value.ToString(),
                    Quantity = (double)rs.Fields.Item("Quantity").Value,
                    InvQuantity = (double)rs.Fields.Item("Quantity").Value,
                    ItemName = rs.Fields.Item("ItemName").Value.ToString(),
                    UOM = rs.Fields.Item("InvntryUom").Value.ToString(),
                    Whse = rs.Fields.Item("Warehouse").Value.ToString(),
                    ReturnReason = rs.Fields.Item("FreeTxt").Value.ToString(),
                    SnBCode = rs.Fields.Item("DistNumber").Value.ToString(),
                    BinCode = rs.Fields.Item("BinCode").Value.ToString()
                }
                );
                rs.MoveNext();
            }

            return lines;
        }
    }
}
