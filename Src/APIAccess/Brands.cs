using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    class Brands : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "Brand";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_BRAND;
        internal const int _sapObjType = (int)SAPbobsCOM.BoObjectTypes.oManufacturers;
        internal const string _keyField = "FirmCode";
        internal const string _nameField = "FirmName";
        internal const string _filterField = _nameField;

        internal override string wmsAPIEndPoint => _wmsAPIEndPoint;
        internal override string wmsObjectType => _objectType;
        internal override int sapObjectType => _sapObjType;
        internal override string sapKeyField => _keyField;
        internal override string sapNameField => _nameField;
        int _firmCode = 0;

        public string FirmCode = "";
        public string FirmName = "";
        public string ObjType = "43";
        public string CreateTS = "";
        public string UpdateTS = "";


        public Brands(int firmCode) : base()
        {
            _firmCode = firmCode;
            FirmCode = _firmCode.ToString();

            var sql = String.Format(Src.Resource.Queries.GET_RECORD_OMRC, _firmCode);
            var rs = eCommon.ExecuteQuery(sql);
            if (rs.RecordCount == 0)
                throw new Exception($"Bran '{_firmCode}' not found [OMRC].");

            FirmName = rs.Fields.Item("FirmName").Value.ToString().Trim();

            CreateTS = DateTime.Now.ToString("yyyyMMdd HH:mm:ss");
            UpdateTS = CreateTS;
        }

        internal override string sapKeyVal => this._firmCode.ToString();
        internal override string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }
    }
}
