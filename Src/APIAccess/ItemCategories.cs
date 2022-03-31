using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    class ItemCategories : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "Category";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_ITEM_CATEGORY;
        internal const int _sapObjType = (int)SAPbobsCOM.BoObjectTypes.oItemGroups;
        internal const string _keyField = "ItmsGrpCod";
        internal const string _nameField = "ItmsGrpNam";
        internal const string _filterField = _nameField;


        internal override string wmsAPIEndPoint => _wmsAPIEndPoint;
        internal override string wmsObjectType => _objectType;
        internal override int sapObjectType => _sapObjType;
        internal override string sapKeyField => _keyField;
        internal override string sapNameField => _nameField;
        private int _itemCategoryCode;

        public int ItemGroupKey = 0;
        public string ItemGroupName = "";
        public string ObjType = "52";
        public string CreateTS;
        public string UpdateTS;

        public ItemCategories(int itemCategoryCode): base()
        {
            _itemCategoryCode = itemCategoryCode;
            ItemGroupKey = _itemCategoryCode;

            var sql = String.Format(Src.Resource.Queries.GET_RECORD_OITB, _itemCategoryCode);
            var rs = eCommon.ExecuteQuery(sql);
            if (rs.RecordCount == 0)
                throw new Exception($"Item Category '{_itemCategoryCode}' not found.");

            ItemGroupName = rs.Fields.Item("ItmsGrpNam").Value.ToString().Trim();

            var CreateDate = (DateTime)rs.Fields.Item("createDate").Value;
            var UpdateDate = (DateTime)rs.Fields.Item("updateDate").Value;
            var CreateTime = 0;
            var UpdateTime = 0;

            CreateTS = eCommon.GetTimeStamp(CreateDate, CreateTime).ToString("yyyyMMdd HH:mm:ss");
            UpdateTS = eCommon.GetTimeStamp(UpdateDate, UpdateTime).ToString("yyyyMMdd HH:mm:ss");

        }

        internal override string sapKeyVal => this._itemCategoryCode.ToString();
        internal override string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }

    }
}
