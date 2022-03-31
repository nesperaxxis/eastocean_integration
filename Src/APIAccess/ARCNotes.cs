using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    class ARCNotes : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "ARCN";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_AR_CN;
        internal const int _sapObjType = (int)SAPbobsCOM.BoObjectTypes.oCreditNotes;
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
        public string DocType = "CREDIT_NOTE";
        public string BaseWMSTransId = "";       //This field is only used when the AR CN is based on DO Return posted from WMS

        public List<RINLine> Lines = new List<RINLine>();

        public ARCNotes(int docEntry) : base()
        {
            _docEntry = docEntry;

            var sql = String.Format(Src.Resource.Queries.GET_RECORD_ORIN, _docEntry);
            var rs = eCommon.ExecuteQuery(sql);
            DocumentKey = (int)rs.Fields.Item("DocEntry").Value;
            DocumentNo = (int)rs.Fields.Item("DocNum").Value;
            DueDate = ((DateTime)rs.Fields.Item("DocDueDate").Value).ToString("yyyyMMdd");
            PostDate = ((DateTime)rs.Fields.Item("DocDate").Value).ToString("yyyyMMdd");
            CardCode = rs.Fields.Item("CardCode").Value.ToString();

            string baseWMSId = rs.Fields.Item(SBOAddon_DB.ODOC_UDF_EXTERNAL_ID).Value.ToString().Trim();
            if (!String.IsNullOrWhiteSpace(baseWMSId))
                BaseWMSTransId = baseWMSId;

            //Lines = RINLine.GetItems(_docEntry);
            Lines = RINLine.GetItems((int)rs.Fields.Item("DocEntry").Value);

        }

        internal override string sapKeyVal => this.DocumentKey.ToString();
        internal override string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }
    }

    public class RINLine
    {
        private static int objType = (int)SAPbobsCOM.BoObjectTypes.oCreditNotes;
        public int DocEntry;
        public int LineNum;
        public string ItemCode;
        public double Quantity;
        public double InvQuantity;
        public string ItemName;
        public string UOM;
        public string Whse = "";
        public string SnBCode = "";              //This field is not used when the AR CN is based on DO Return posted from WMS
        public string BinCode = "";              //This field is not used when the AR CN is based on DO Return posted from WMS
        public string ReturnReason = "";
        public string BaseWMSTransId = "";       //This field is only used when the AR CN is based on DO Return posted from WMS

        public RINLine() { }

        public static List<RINLine> GetItems(int docEntry)
        {
            List<RINLine> lines = new List<RINLine>();
            string tableName = "RIN"; //eCommon.GetTableName(objType.ToString());
            string sql = String.Format(Resource.Queries.OINM_GET_LINE_SNB_BIN_INFO, 14, docEntry, tableName);
            var rs = eCommon.ExecuteQuery(sql);
            for (int i = 0; i < rs.RecordCount; i++)
            {
                var rinLine = new RINLine
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
                };

                string snbCode = String.IsNullOrEmpty(rs.Fields.Item("DistNumber").Value.ToString()) ? "" : rs.Fields.Item("DistNumber").Value.ToString().Trim();
                string binCode = String.IsNullOrEmpty(rs.Fields.Item("BinCode").Value.ToString()) ? "" : rs.Fields.Item("BinCode").Value.ToString().Trim();
                string baseWMSId = rs.Fields.Item(SBOAddon_DB.DOC1_UDF_EXTERNAL_ID).Value.ToString().Trim();
                if (snbCode != null)
                    rinLine.SnBCode = snbCode;
                if (binCode != null)
                    rinLine.BinCode = binCode;
                if (baseWMSId != null)
                    rinLine.BaseWMSTransId = baseWMSId;

                lines.Add(rinLine);
                rs.MoveNext();
            }

            return lines;
        }
    }
}
