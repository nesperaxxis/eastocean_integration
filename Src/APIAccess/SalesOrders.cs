using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    class SalesOrders : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "SalesOrder";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_SALES_ORDER;
        internal const int _sapObjType = (int)SAPbobsCOM.BoObjectTypes.oOrders;
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
        public string CardName;
        public string CustRefNum = "";
        public string SlpName = "";
        public string Location = "";
        public string VehicleNumber = "";
        public string Remark = "";
        public string ShipToCode = "";
        public string DocumentStatus = "";
        public string Canceled = "";
        public string DirectDelivery = "";

        public List<SOLines> Lines = new List<SOLines>();

        public SalesOrders(int docEntry) : base()
        {
            _docEntry = docEntry;

            var sql = String.Format(Src.Resource.Queries.GET_RECORD_ORDR, _docEntry);
            var rs = eCommon.ExecuteQuery(sql);
            if (rs.RecordCount == 0)
                throw new Exception($"Sales Order not found. [ORDR.DocEntry]={docEntry}");
            DocumentKey = (int)rs.Fields.Item("DocEntry").Value;
            DocumentNo = (int)rs.Fields.Item("DocNum").Value;
            PostDate = ((DateTime)rs.Fields.Item("DocDate").Value).ToString("yyyyMMdd");
            DeliveryDate = ((DateTime)rs.Fields.Item("DocDueDate").Value).ToString("yyyyMMdd");
            CardCode = rs.Fields.Item("CardCode").Value.ToString();
            CardName = rs.Fields.Item("CardName").Value.ToString();
            CustRefNum = rs.Fields.Item("NumAtCard").Value.ToString();
            SlpName = rs.Fields.Item("SlpName").Value.ToString();
            Remark = rs.Fields.Item("Comments").Value.ToString() + (rs.Fields.Item("CANCELED").Value.ToString() == "Y" ? " - CANCELED" : "");
            ShipToCode = rs.Fields.Item("ShipToCode").Value.ToString();
            DocumentStatus = rs.Fields.Item("DocStatus").Value.ToString();
            Canceled = rs.Fields.Item("CANCELED").Value.ToString();
            DirectDelivery = rs.Fields.Item("DirectDelivery").Value.ToString();

            Lines = SOLines.GetItems(_docEntry);
        }

        internal override string sapKeyVal => this.DocumentKey.ToString();

        internal override string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }
    }

    public class SOLines
    {

        public int DocEntry;
        public int LineNum;
        public string ItemCode = "";
        public string UOM = "";
        public double Quantity = 0;
        public double InvQuantity = 0;
        public string ItemName = "";
        public string FreeText = "";
        public string Whse = "";

        public SOLines() { }

        public static List<SOLines> GetItems(int docEntry)
        {
            List<SOLines> lines = new List<SOLines>();
            string sql = String.Format(Resource.Queries.GET_RECORD_RDR1, docEntry);
            var rs = eCommon.ExecuteQuery(sql);
            for (int i = 0; i < rs.RecordCount; i++)
            {
                lines.Add(new SOLines
                {
                    DocEntry = (int)rs.Fields.Item("DocEntry").Value,
                    LineNum = (int)rs.Fields.Item("LineNum").Value,
                    ItemCode = rs.Fields.Item("ItemCode").Value.ToString(),
                    Quantity = (double)rs.Fields.Item("OpenQty").Value,
                    InvQuantity = (double)rs.Fields.Item("OpenInvQty").Value,
                    ItemName = rs.Fields.Item("Dscription").Value.ToString(),
                    UOM = rs.Fields.Item("unitMsr").Value.ToString(),
                    Whse = rs.Fields.Item("WhsCode").Value.ToString(),
                    FreeText = rs.Fields.Item("FreeTxt").Value.ToString()
                }
                );
                rs.MoveNext();
            }

            return lines;
        }
    }
}
