using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    class Warehouses : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "Warehouse";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_WAREHOUSE;
        internal const int _sapObjType = (int)SAPbobsCOM.BoObjectTypes.oWarehouses;
        internal const string _keyField = "WhsCode";
        internal const string _nameField = "WhsName";
        internal const string _filterField = _keyField;


        internal override string wmsAPIEndPoint => _wmsAPIEndPoint;
        internal override string wmsObjectType => _objectType;
        internal override int sapObjectType => _sapObjType;
        internal override string sapKeyField => _keyField;
        internal override string sapNameField => _nameField;
        string _whsCode = "";

        public string ObjType = "64";
        public string WhsCode = "";
        public string WhsName = "";
        public string WhsFName = "";
        public string Address1 = "";
        public string Address2 = "";
        public string PostalCode = "";
        public string City = "";
        public string Phone = "";
        public string Contact = "";
        public string FaxNumber = "";
        public string Email = "";
        public string Website = "";
        public string BinEnabled = "";
        public string CreateTS = "";
        public string UpdateTS = "";

        public Warehouses(string whsCode) : base()
        {
            _whsCode = whsCode;
            WhsCode = _whsCode;

            var sql = String.Format(Src.Resource.Queries.GET_RECORD_OWHS, _whsCode);
            var rs = eCommon.ExecuteQuery(sql);
            if (rs.RecordCount == 0)
                throw new Exception($"Warehouse '{_whsCode}' not found.");
            WhsName = rs.Fields.Item("WhsName").Value.ToString().Trim();
            Address1 = rs.Fields.Item("Street").Value.ToString().Trim();
            Address2 = rs.Fields.Item("Block").Value.ToString().Trim() + rs.Fields.Item("Building").Value.ToString().Trim();
            PostalCode = rs.Fields.Item("ZipCode").Value.ToString().Trim();
            City = rs.Fields.Item("City").Value.ToString().Trim();
            BinEnabled = rs.Fields.Item("BinActivat").Value.ToString().Trim();

            var CreateDate = (DateTime)rs.Fields.Item("createDate").Value;
            var UpdateDate = (DateTime)rs.Fields.Item("updateDate").Value;
            var CreateTime = 0;
            var UpdateTime = 0;

            CreateTS = eCommon.GetTimeStamp(CreateDate, CreateTime).ToString("yyyyMMdd HH:mm:ss");
            UpdateTS = eCommon.GetTimeStamp(UpdateDate, UpdateTime).ToString("yyyyMMdd HH:mm:ss");
        }

        internal override string sapKeyVal => this.WhsCode.ToString();
        internal override string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }
    }
}
