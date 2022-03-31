using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    class BinLocations : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "WarehouseBin";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_BIN;
        internal const int _sapObjType = 10000206;
        internal const string _keyField = "AbsEntry";
        internal const string _nameField = "BinCode";
        internal const string _filterField = _nameField;


        internal override string wmsAPIEndPoint => _wmsAPIEndPoint;
        internal override string wmsObjectType => _objectType;
        internal override int sapObjectType => _sapObjType;
        internal override string sapKeyField => _keyField;
        internal override string sapNameField => _nameField;
        int _binEntry = 0;

        public int SAPBinAbs = 0;
        public string BinCode = "";
        public string BinName = "";
        public string ObjType = _sapObjType.ToString();
        public string WhsCode = "";
        public string CreateTS = "";
        public string UpdateTS = "";

        public BinLocations(int binEntry) : base()
        {
            _binEntry = binEntry;
            SAPBinAbs = _binEntry;

            var sql = String.Format(Src.Resource.Queries.GET_RECORD_OBIN, _binEntry);
            var rs = eCommon.ExecuteQuery(sql);
            BinCode = rs.Fields.Item("BinCode").Value.ToString().Trim();
            BinName = BinCode;
            WhsCode = rs.Fields.Item("WhsCode").Value.ToString().Trim();
            CreateTS = ((DateTime)rs.Fields.Item("CreateDate").Value).ToString("yyyyMMdd HH:mm:ss");
            UpdateTS = ((DateTime)rs.Fields.Item("UpdateDate").Value).ToString("yyyyMMdd HH:mm:ss");
        }

        internal override string sapKeyVal => this._binEntry.ToString();
        internal override string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }
    }
}