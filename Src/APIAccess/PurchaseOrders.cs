using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    class PurchaseOrders : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "PurchaseOrder";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_PURCHASE_ORDER;
        internal const int _sapObjType = (int)SAPbobsCOM.BoObjectTypes.oPurchaseOrders;
        internal const string _keyField = "DocEntry";
        internal const string _nameField = "DocNum";
        internal const string _filterField = _nameField;


        internal override string wmsAPIEndPoint => _wmsAPIEndPoint;
        internal override string wmsObjectType => _objectType;
        internal override int sapObjectType => _sapObjType;
        internal override string sapKeyField => _keyField;
        internal override string sapNameField => _nameField;
        int _docEntry = 0;

        public string ObjType = _sapObjType.ToString();
        public int DocumentKey;
        public int DocumentNo;
        public string PostDate;
        public string DeliveryDate;
        public string CardCode;
        public string SlpName = "";
        public string Location = "";
        public string Remark = "";
        public string PayTerm = "";
        public string Currency = "";
        public double ExcRate = 0;
        public string Canceled = "";
        public string Status = "";


        public List<POLines> Lines = new List<POLines>();
        public PurchaseOrders(int docEntry) : base()
        {
            _docEntry = docEntry;

            var sql = String.Format(Src.Resource.Queries.GET_RECORD_OPOR, _docEntry);
            var rs = eCommon.ExecuteQuery(sql);
            if (rs.RecordCount == 0)
                throw new Exception($"Purchase Order not found. [OPOR.DocEntry]={docEntry}");
            DocumentKey = (int)rs.Fields.Item("DocEntry").Value;
            DocumentNo = (int)rs.Fields.Item("DocNum").Value;
            PostDate = ((DateTime)rs.Fields.Item("DocDate").Value).ToString("yyyyMMdd");
            DeliveryDate = ((DateTime)rs.Fields.Item("DocDueDate").Value).ToString("yyyyMMdd");
            CardCode = rs.Fields.Item("CardCode").Value.ToString();
            SlpName = rs.Fields.Item("SlpName").Value.ToString();
            Remark = rs.Fields.Item("Comments").Value.ToString() + (rs.Fields.Item("CANCELED").Value.ToString() == "Y" ? " - CANCELED" : "");
            Currency = rs.Fields.Item("DocCur").Value.ToString();
            ExcRate = (Double)rs.Fields.Item("DocRate").Value;
            Canceled = rs.Fields.Item("CANCELED").Value.ToString();
            Status = rs.Fields.Item("DocStatus").Value.ToString();
            Lines = POLines.GetItems(_docEntry);
        }

        internal override string sapKeyVal => this.DocumentKey.ToString();

        internal override string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }
    }

    public class POLines
    {

        public int DocEntry;
        public int LineNum;
        public string ItemCode;
        public string UOM;
        public double Quantity;
        public double InvQuantity;
        public string ItemName;
        public string FreeText;
        public string Whse;

        public POLines() { }

        public static List<POLines> GetItems(int docEntry)
        {
            List<POLines> lines = new List<POLines>();
            string sql = String.Format(Resource.Queries.GET_RECORD_POR1, docEntry);
            var rs = eCommon.ExecuteQuery(sql);
            for (int i = 0; i < rs.RecordCount; i++)
            {
                lines.Add(new POLines
                {
                    DocEntry = (int)rs.Fields.Item("DocEntry").Value,
                    LineNum = (int)rs.Fields.Item("LineNum").Value,
                    ItemCode = rs.Fields.Item("ItemCode").Value.ToString(),
                    Quantity = (double)rs.Fields.Item("OpenQty").Value,
                    InvQuantity = (double)rs.Fields.Item("OpenInvQty").Value,
                    ItemName = rs.Fields.Item("Dscription").Value.ToString(),
                    UOM = rs.Fields.Item("unitMsr").Value.ToString(),
                    Whse = rs.Fields.Item("WhsCode").Value.ToString(),
                }
                );
                rs.MoveNext();
            }

            return lines;
        }
    }
}
