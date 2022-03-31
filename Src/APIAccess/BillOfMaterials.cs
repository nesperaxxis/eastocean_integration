using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    class BillOfMaterials : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "BOM";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_BOM;
        internal const int _sapObjType = (int)SAPbobsCOM.BoObjectTypes.oProductTrees;
        internal const string _keyField = "Code";
        internal const string _nameField = "Name";
        internal const string _filterField =_keyField;


        internal override string wmsAPIEndPoint => _wmsAPIEndPoint;
        internal override string wmsObjectType => _objectType;
        internal override int sapObjectType => _sapObjType;
        internal override string sapKeyField => _keyField;
        internal override string sapNameField => _nameField;
        string _itemCode = "";

        public string ObjType = _sapObjType.ToString();
        public string FGItemCode = "";
        public string ItemName = "";
        public double PlanQty;
        public string Whse = "";
        public string CreateTS = "";
        public string UpdateTS = "";

        public List<Component> Lines = new List<Component>();


        public BillOfMaterials(string itemCode) : base()
        {
            _itemCode = itemCode;
            FGItemCode = itemCode;

            var sql = String.Format(Src.Resource.Queries.GET_RECORD_OITT, _itemCode);
            var rs = eCommon.ExecuteQuery(sql);
            ItemName = rs.Fields.Item("Name").Value.ToString().Trim();
            PlanQty = (double)rs.Fields.Item("Qauntity").Value;
            Whse = rs.Fields.Item("ToWH").Value.ToString().Trim();

            var CreateDate = (DateTime)rs.Fields.Item("CreateDate").Value;
            var UpdateDate = (DateTime)rs.Fields.Item("UpdateDate").Value;
            var CreateTime = (int)rs.Fields.Item("CreateTS").Value;
            var UpdateTime = (int)rs.Fields.Item("UpdateTS").Value;

            CreateTS = eCommon.GetTimeStamp(CreateDate, CreateTime).ToString("yyyyMMdd HH:mm:ss");
            UpdateTS = eCommon.GetTimeStamp(UpdateDate, UpdateTime).ToString("yyyyMMdd HH:mm:ss");

            Lines = Component.GetComponents(_itemCode);
        }

        internal override string sapKeyVal => this._itemCode.ToString();
        internal override string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }
    }

    public class Component
    {
        public string ItemCode;
        public string ItemName;
        public double Quantity;
        public string UOM;
        public string Whse;


        public Component() { }

        public static List<Component> GetComponents(string itemCode)
        {
            List<Component> addresses = new List<Component>();
            string sql = String.Format(Resource.Queries.GET_RECORD_ITT1_ITEMS, itemCode.Replace("'", "''"));
            var rs = eCommon.ExecuteQuery(sql);
            for (int i = 0; i < rs.RecordCount; i++)
            {
                addresses.Add(new Component
                {
                    ItemCode = rs.Fields.Item("Code").Value.ToString(),
                    Quantity = (double)rs.Fields.Item("Quantity").Value,
                    ItemName = rs.Fields.Item("ItemName").Value.ToString(),
                    UOM = String.IsNullOrEmpty(rs.Fields.Item("Uom").Value.ToString()) ? rs.Fields.Item("iUom").Value.ToString() : rs.Fields.Item("Uom").Value.ToString(),
                    Whse = rs.Fields.Item("Warehouse").Value.ToString()
                }
                );
                rs.MoveNext();
            }

            return addresses;
        }

    }
}
