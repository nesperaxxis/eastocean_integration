using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    class WorkOrders : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "ProductionOrder";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_WORK_ORDER;
        internal const int _sapObjType = (int)SAPbobsCOM.BoObjectTypes.oProductionOrders;
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
        public string PostDate = "";
        public string DueDate = "";
        public string Status = "";
        public string AgentID = "";
        public string WhsCode = "";
        public string FGItemCode = "";
        public double PlannedQty = 0;
        public string ProductionType = "S";     //S = Standard; D = Disassembly

        public List<WORLine> Lines = new List<WORLine>();


        public WorkOrders(int docEntry) : base()
        {
            _docEntry = docEntry;

            var sql = String.Format(Src.Resource.Queries.GET_RECORD_OWOR, _docEntry);
            var rs = eCommon.ExecuteQuery(sql);
            DocumentKey = (int)rs.Fields.Item("DocEntry").Value;
            DocumentNo = (int)rs.Fields.Item("DocNum").Value;
            PostDate = ((DateTime)rs.Fields.Item("PostDate").Value).ToString("yyyyMMdd");
            DueDate = ((DateTime)rs.Fields.Item("DueDate").Value).ToString("yyyyMMdd");
            Status = rs.Fields.Item("Status").Value.ToString().Trim();
            WhsCode = rs.Fields.Item("Warehouse").Value.ToString().Trim();
            FGItemCode = rs.Fields.Item("ItemCode").Value.ToString().Trim();
            PlannedQty = (double)rs.Fields.Item("PlannedQty").Value;
            ProductionType = rs.Fields.Item("Type").Value.ToString().Trim();

            Lines = WORLine.GetItems(_docEntry);

        }

        internal override string sapKeyVal => this.DocumentKey.ToString();
        internal override string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }
    }

    public class WORLine
    {
        public int DocEntry;
        public int LineNum;
        public string ItemCode;
        public string UOM;
        public double BaseQuantity = 0;
        public double PlannedQuantity = 0;
        public string ItemName;
        public string Whse;
        public string Type;

        public WORLine() { }

        public static List<WORLine> GetItems(int docEntry)
        {
            List<WORLine> lines = new List<WORLine>();
            string sql = String.Format(Resource.Queries.GET_RECORD_WOR1, docEntry);
            var rs = eCommon.ExecuteQuery(sql);
            for (int i = 0; i < rs.RecordCount; i++)
            {
                lines.Add(new WORLine
                {
                    DocEntry = (int)rs.Fields.Item("DocEntry").Value,
                    LineNum = (int)rs.Fields.Item("LineNum").Value,
                    ItemCode = rs.Fields.Item("ItemCode").Value.ToString(),
                    BaseQuantity = (double)rs.Fields.Item("BaseQty").Value,
                    PlannedQuantity = (double)rs.Fields.Item("PlannedQty").Value,
                    ItemName = rs.Fields.Item("ItemName").Value.ToString(),
                    UOM = rs.Fields.Item("InvntryUom").Value.ToString(),
                    Whse = rs.Fields.Item("wareHouse").Value.ToString(),
                    Type = "ITEM"
                }
                ); ; ;
                rs.MoveNext();
            }

            return lines;
        }
    }
}
