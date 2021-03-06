using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    class ARResInvoices : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "ReserveInvoice";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_RESERVE_INVOICE;
        internal const int _sapObjType = (int)SAPbobsCOM.BoObjectTypes.oInvoices;
        internal const string _keyField = "DocEntry";
        internal const string _nameField = "DocNum";
        internal const string _filterField = _nameField;


        internal override string wmsAPIEndPoint => _wmsAPIEndPoint;
        internal override string wmsObjectType => _objectType;
        internal override int sapObjectType => _sapObjType;
        internal override string sapKeyField => _keyField;
        internal override string sapNameField => _nameField;
        int _docEntry = 0;

        public string ObjType = "13R";
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
        public string DirectDelivery = "";

        public List<RILines> Lines = new List<RILines>();
        public ARResInvoices(int docEntry) : base()
        {
            _docEntry = docEntry;

            var sql = String.Format(Src.Resource.Queries.GET_RECORD_OINV, _docEntry);
            var rs = eCommon.ExecuteQuery(sql);
            if (rs.RecordCount == 0)
                throw new Exception($"AR Reserve Invoice not found. [OINV.DocEntry]={docEntry}");
            DocumentKey = (int)rs.Fields.Item("DocEntry").Value;
            DocumentNo = (int)rs.Fields.Item("DocNum").Value;
            PostDate = ((DateTime)rs.Fields.Item("DocDate").Value).ToString("yyyyMMdd");
            DeliveryDate = ((DateTime)rs.Fields.Item("DocDueDate").Value).ToString("yyyyMMdd");
            CardCode = rs.Fields.Item("CardCode").Value.ToString();
            CardName = rs.Fields.Item("CardName").Value.ToString();
            CustRefNum = rs.Fields.Item("NumAtCard").Value.ToString();
            SlpName = rs.Fields.Item("SlpName").Value.ToString();
            Remark = rs.Fields.Item("Comments").Value.ToString() + (rs.Fields.Item("CANCELED").Value.ToString()=="Y" ? " - CANCELED" : "");
            ShipToCode = rs.Fields.Item("ShipToCode").Value.ToString();
            DirectDelivery = rs.Fields.Item("DirectDelivery").Value.ToString();

            Lines = RILines.GetItems(_docEntry);
        }

        internal override string sapKeyVal => this.DocumentKey.ToString();
        internal override string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }
    }

    public class RILines
    {

        public int DocEntry;
        public int LineNum;
        public string ItemCode;
        public string UOM;
        public double Quantity;
        public double InvQuantity;
        public string ItemName;
        public string FreeText = "";
        public string Whse = "";

        public RILines() { }

        public static List<RILines> GetItems(int docEntry)
        {
            List<RILines> lines = new List<RILines>();
            string sql = String.Format(Resource.Queries.GET_RECORD_INV1, docEntry);
            var rs = eCommon.ExecuteQuery(sql);
            for (int i = 0; i < rs.RecordCount; i++)
            {
                lines.Add(new RILines
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
