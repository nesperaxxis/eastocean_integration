using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.IO;


namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    class Items : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "Item";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_ITEM;
        internal const int _sapObjType = (int)SAPbobsCOM.BoObjectTypes.oItems;
        internal const string _keyField = "ItemCode";
        internal const string _nameField = "ItemName";
        internal const string _filterField = _keyField;


        internal override string wmsAPIEndPoint => _wmsAPIEndPoint;
        internal override string wmsObjectType => _objectType;
        internal override int sapObjectType => _sapObjType;
        internal override string sapKeyField => _keyField;
        internal override string sapNameField => _nameField;
        private string _itemCode = "";

        public string ItemCode = "";
        public string ItemName = "";
        public string ObjType = "4";
        public string FrgnName = "";
        public string ShortName = "";
        public string InventoryUOM = "";
        public string SalesUOM = "";
        public string PurchaseUOM = "";
        public string VendorCode = "";
        public int ItemGroupKey;
        public string ItemGroupName = "";
        public string Brand = "";
        public string LotTracking = "";
        public string Active = "";
        public string SerialTracking = "";
        public string DefWhse = "";
        public string DefBin = "";
        public string MinShelfLife = "";
        public double PchLength;
        public string PchLengthUnit = "";
        public double PchWidth;
        public string PchWidthUnit = "";
        public double PchHeight;
        public string PchHeightUnit = "";
        public double PchVolume;
        public string PchVolumeUnit = "";
        public string ProductRanking = "";
        public string PictureName = "";
        public string PictureData = "";  //Base64 = "";
        public string CreateTS = "";
        public string UpdateTS = "";


        public Items(string itemCode) : base()
        {
            _itemCode = itemCode;
            ItemCode = _itemCode;

            var sql = String.Format(Src.Resource.Queries.GET_RECORD_OITM, _itemCode);
            var rs = eCommon.ExecuteQuery(sql);
            ItemName = rs.Fields.Item("ItemName").Value.ToString().Trim();
            FrgnName = String.IsNullOrEmpty(rs.Fields.Item("FrgnName").Value.ToString().Trim()) ? ItemName : rs.Fields.Item("FrgnName").Value.ToString().Trim();
            ShortName = rs.Fields.Item("ItemName").Value.ToString().Trim();
            InventoryUOM = rs.Fields.Item("InvntryUom").Value.ToString().Trim();
            SalesUOM = rs.Fields.Item("SalUnitMsr").Value.ToString().Trim();
            PurchaseUOM = rs.Fields.Item("BuyUnitMsr").Value.ToString().Trim();
            VendorCode = rs.Fields.Item("CardCode").Value.ToString().Trim();
            ItemGroupKey = (int)rs.Fields.Item("ItmsGrpCod").Value;
            ItemGroupName = rs.Fields.Item("ItmsGrpNam").Value.ToString().Trim();
            Brand = rs.Fields.Item("FrgnName").Value.ToString().Trim();
            LotTracking = rs.Fields.Item("ManBtchNum").Value.ToString().Trim();
            SerialTracking = rs.Fields.Item("ManSerNum").Value.ToString().Trim();
            DefWhse = rs.Fields.Item("DfltWH").Value.ToString().Trim();
            DefBin = rs.Fields.Item("BinCode").Value.ToString().Trim();
            PchLength = (double)rs.Fields.Item("BLength1").Value;
            PchLengthUnit = eCommon.GetLengthUnit((int)rs.Fields.Item("BLen1Unit").Value);
            PchWidth = (double)rs.Fields.Item("BWidth1").Value;
            PchWidthUnit = eCommon.GetLengthUnit((int)rs.Fields.Item("BWdth1Unit").Value);
            PchHeight = (double)rs.Fields.Item("BHeight1").Value;
            PchHeightUnit = eCommon.GetLengthUnit((int)rs.Fields.Item("BHght1Unit").Value);
            PchVolume = (double)rs.Fields.Item("BVolume").Value;
            PchVolumeUnit = eCommon.GetVolumeUnit((int)rs.Fields.Item("BVolUnit").Value);
            ProductRanking = rs.Fields.Item("U_PRANK").Value.ToString().Trim();
            var picturePath = rs.Fields.Item("BitmapPath").Value.ToString().Trim();
            //PictureName = String.IsNullOrEmpty(rs.Fields.Item("PicturName").Value.ToString().Trim()) ? "noimage.jpg" : rs.Fields.Item("PicturName").Value.ToString().Trim();
            PictureName = String.IsNullOrEmpty(rs.Fields.Item("PicturName").Value.ToString().Trim()) ? "" : rs.Fields.Item("PicturName").Value.ToString().Trim();
            PictureData = String.IsNullOrEmpty(PictureName) ? "" : PicConversion.CreateBase64Image(PictureName);
             //String.IsNullOrEmpty(eCommon.GetPictureContent(picturePath, PictureName)) ? "" : eCommon.GetPictureContent(picturePath, PictureName);  //Base64 = "";
            //if (String.IsNullOrWhiteSpace(PictureData)) PictureName = null;
            Active = rs.Fields.Item("validFor").Value.ToString() == "Y" || (rs.Fields.Item("validFor").Value.ToString() == "N" && rs.Fields.Item("frozenFor").Value.ToString() == "N") ? "Y" : "N";

            var CreateDate = (DateTime)rs.Fields.Item("CreateDate").Value;
            var UpdateDate = (DateTime)rs.Fields.Item("UpdateDate").Value;
            var CreateTime = (int)rs.Fields.Item("CreateTS").Value;
            var UpdateTime = (int)rs.Fields.Item("CreateTS").Value;

            CreateTS = eCommon.GetTimeStamp(CreateDate, CreateTime).ToString("yyyyMMdd HH:mm:ss");
            UpdateTS = eCommon.GetTimeStamp(UpdateDate, UpdateTime).ToString("yyyyMMdd HH:mm:ss");

        }

        internal override string sapKeyVal => this._itemCode.ToString();
        internal override string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }

    }
    public class PicConversion
    {
        public static string CreateBase64Image(string strPicName)
        {
            //image to byteArray
            Image img = Image.FromFile(@"C:\Program Files (x86)\SAP\SAP Business One\Bitmaps\" + strPicName);
            byte[] bArr = imgToByteArray(img);

            //byte[] bArr = imgToByteConverter(img);

            //Again convert byteArray to image and displayed in a picturebox
            //Image img1 = byteArrayToImage(bArr);

            return Convert.ToBase64String(bArr);

        }
        //convert image to bytearray
        public static byte[] imgToByteArray(Image img)
        {
            using (MemoryStream mStream = new MemoryStream())
            {
                img.Save(mStream, img.RawFormat);
                return mStream.ToArray();
            }
        }
        //convert bytearray to image
        public static Image byteArrayToImage(byte[] byteArrayIn)
        {
            using (MemoryStream mStream = new MemoryStream(byteArrayIn))
            {
                return Image.FromStream(mStream);
            }
        }
        //another easy way to convert image to bytearray
        public static byte[] imgToByteConverter(Image inImg)
        {
            ImageConverter imgCon = new ImageConverter();
            return (byte[])imgCon.ConvertTo(inImg, typeof(byte[]));
        }
    }
}
