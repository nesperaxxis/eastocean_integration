using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    class Barcodes : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "Barcode";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_BAR_CODE;
        internal const int _sapObjType = 1470000062;
        internal const string _keyField = "BcdEntry";
        internal const string _nameField = "BcdCode";
        internal const string _filterField = _nameField;


        internal override string wmsAPIEndPoint => _wmsAPIEndPoint;
        internal override string wmsObjectType => _objectType;
        internal override int sapObjectType => _sapObjType;
        internal override string sapKeyField => _keyField;
        internal override string sapNameField => _nameField;
        string _itemCode = "";
        int _bcdEntry = 0;

        public string ObjType = _sapObjType.ToString();
        public string ItemCode = "";
        public List<BarCodeObject> barcodes = new List<BarCodeObject>();

        public Barcodes(string itemCode) : base()
        {
            _itemCode = itemCode;
            ItemCode = _itemCode;

            string sql = String.Format( Src.Resource.Queries.OITM_GET_BARCODES, itemCode.Replace("'","''"));
            var rs = eCommon.ExecuteQuery(sql);
            for(int i =0; i<rs.RecordCount; i++)
            {
                barcodes.Add(new BarCodeObject(itemCode,
                    (int)rs.Fields.Item("BcdEntry").Value,
                    rs.Fields.Item("BcdCode").Value.ToString(),
                    rs.Fields.Item("BcdName").Value.ToString(),
                    rs.Fields.Item("UomCode").Value.ToString(),
                    (DateTime)rs.Fields.Item("CreateDate").Value,
                    (DateTime)rs.Fields.Item("UpdateDate").Value));

                rs.MoveNext();
            }
        }

        public Barcodes(int bcdEntry) : base()
        {
            _bcdEntry = bcdEntry;

            string sql = String.Format(Src.Resource.Queries.GET_RECORD_OBCD, _bcdEntry);
            var rs = eCommon.ExecuteQuery(sql);
            for (int i = 0; i < rs.RecordCount; i++)
            {
                if (String.IsNullOrEmpty(_itemCode))
                {
                    _itemCode = rs.Fields.Item("ItemCode").Value.ToString();
                    ItemCode = _itemCode;
                }

                barcodes.Add(new BarCodeObject(_itemCode, _bcdEntry,
                    rs.Fields.Item("BcdCode").Value.ToString(),
                    rs.Fields.Item("BcdName").Value.ToString(),
                    rs.Fields.Item("UomCode").Value.ToString(),
                    (DateTime)rs.Fields.Item("CreateDate").Value,
                    (DateTime)rs.Fields.Item("UpdateDate").Value)) ;

                rs.MoveNext();
            }
        }

        internal override string sapKeyVal => this._itemCode.ToString();
        internal override string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }
    }

    class BarCodeObject
    {
        public int SAPBcdEntry = 0;
        public string ItemCode = "";
        public string BarCode = "";
        public string BarCodeDescription = "";
        public string UOM = "";
        public string CreateTS { get { return _createDate.ToString("yyyyMMdd HH:mm:ss"); } }
        public string UpdateTS { get { return _updateDate.ToString("yyyyMMdd HH:mm:ss"); } }


        private DateTime _createDate; 
        private DateTime _updateDate; 

        public BarCodeObject(string itemCode, int bcdEntry, string barCode, string bcdName, string uom, DateTime createDate, DateTime updateDate )
        {
            ItemCode = itemCode;
            SAPBcdEntry = bcdEntry;
            BarCode = barCode;
            BarCodeDescription = bcdName;
            UOM = uom;
            _createDate = createDate;
            _updateDate = updateDate;
        }
    }
}
