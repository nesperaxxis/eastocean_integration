using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    class APReturns : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "APGoodsReturn";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_AP_RETURN;
        internal const int _sapObjType = (int)SAPbobsCOM.BoObjectTypes.oPurchaseReturns;    
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
        public string CardCode;
        public string DueDate;
        public string DocType;
        public int BasePOKey = 0;
        public int BasePONo = 0;

        List<RPDLine> Lines = new List<RPDLine>();


        public APReturns(int docEntry) : base()
        {
            _docEntry = docEntry;

            var sql = String.Format(Src.Resource.Queries.GET_RECORD_ORPD, _docEntry);
            var rs = eCommon.ExecuteQuery(sql);
            DocumentKey = (int)rs.Fields.Item("DocEntry").Value;
            DocumentNo = (int)rs.Fields.Item("DocNum").Value;
            PostDate = ((DateTime)rs.Fields.Item("DocDate").Value).ToString("yyyyMMdd");
            CardCode = rs.Fields.Item("CardCode").Value.ToString();
            DueDate = ((DateTime)rs.Fields.Item("DocDueDate").Value).ToString("yyyyMMdd");
            DocType = "Goods Returns";
            if((int)rs.Fields.Item("BaseEntry").Value != 0) {
                BasePOKey = (int)rs.Fields.Item("BaseEntry").Value;
                int.TryParse(rs.Fields.Item("BaseRef").Value.ToString(), out BasePONo);
            }
            Lines = RPDLine.GetItems(_docEntry);
        }

        internal override string sapKeyVal => this.DocumentKey.ToString();
        internal override string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }
    }

    public class RPDLine
    {
        private static int objType = (int)SAPbobsCOM.BoObjectTypes.oPurchaseReturns;
        public int DocEntry;
        public int LineNum;
        public string ItemCode;   
        public double Quantity;
        public double InvQuantity;
        public string ItemName;
        public string UOM;
        public string Whse = "";
        public string SnBCode ="";
        public string BinCode = "";
        public string ReturnReason = "";

        public RPDLine() { }

        public static List<RPDLine> GetItems(int docEntry)
        {
            List<RPDLine> lines = new List<RPDLine>();
            string tableName = "RPD";//eCommon.GetTableName(objType.ToString());
            string sql = String.Format(Resource.Queries.OINM_GET_LINE_SNB_BIN_INFO, 21, docEntry, tableName);

            var rs = eCommon.ExecuteQuery(sql);
            for (int i = 0; i < rs.RecordCount; i++)
            {
                lines.Add(new RPDLine
                {
                    DocEntry = (int)rs.Fields.Item("DocEntry").Value,
                    LineNum = (int)rs.Fields.Item("LineKey").Value,
                    ItemCode = rs.Fields.Item("ItemCode").Value.ToString(),
                    Quantity = (double)rs.Fields.Item("Quantity").Value,
                    InvQuantity = (double)rs.Fields.Item("Quantity").Value,
                    ItemName = rs.Fields.Item("ItemName").Value.ToString(),
                    UOM = rs.Fields.Item("InvntryUom").Value.ToString(),
                    Whse = rs.Fields.Item("Warehouse").Value.ToString(),
                    SnBCode = String.IsNullOrEmpty(rs.Fields.Item("DistNumber").Value.ToString()) ? "" : rs.Fields.Item("DistNumber").Value.ToString(),
                    BinCode = String.IsNullOrEmpty(rs.Fields.Item("BinCode").Value.ToString()) ? "" : rs.Fields.Item("BinCode").Value.ToString(),
                    ReturnReason = rs.Fields.Item("FreeTxt").Value.ToString()
                }
                );
                rs.MoveNext();
            }

            return lines;
        }
        
    }
}
