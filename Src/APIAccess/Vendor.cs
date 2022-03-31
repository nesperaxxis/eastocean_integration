using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    internal class Vendor : WMSSynchObject
    {
        internal const string _wmsAPIEndPoint = "Vendor";        //TODO: Provide the Simplr WMS API EndPoint to call here, excluding the BaseAddress
        internal const string _objectType = SBOAddon.SYNCH_O_OBJECT_VENDOR;
        internal const int _sapObjType = (int)SAPbobsCOM.BoObjectTypes.oBusinessPartners;
        internal const string _keyField = "CardCode";
        internal const string _nameField = "CardName";
        internal const string _filterField = _keyField;



        internal override string wmsAPIEndPoint => _wmsAPIEndPoint;
        internal override string wmsObjectType => _objectType;
        internal override int sapObjectType => _sapObjType;
        internal override string sapKeyField => _keyField;
        internal override string sapNameField => _nameField;
        private string _CardCode = "";
        public String ObjType = "";
        public String CardCode = "";     //Required BP code
        public String CardName = "";       //Required    BP Name
        public String ForeignName = "";        //Optional BP Foreign Name
        public String ShortName = "";          //Optional BP Name
        public String BillTo = "";              //Required
        public String BillToStreet = "";            //Optional    Bill-to Street
        public String BillToBlock = "";       //Optional Bill-to Block
        public String BillToPostCode = "";         //Optional Bill-to Postcode
        public String BillToCity = "";               //Optional    Bill-to City
        public String BillToCountry = "";            //Optional    Bill-to Country
        public String Telephone = "";              //Optional    Telephone 1
        public String ContactPerson = "";      //Optional Contact Person
        public String Active = "";     //Required    Active(Y/N) change to 1/0 required by Simplr
        public String ZoneCode = "";           //Optional N/A
        public String FaxNumber = "";              //Fax Number
        public String Email = "";              //Optional    E-Mail
        public String WebSite = "";            //Optional Web Site
        public String ShipmentMethod = "";     //Optional    N/A
        public int SalesEmployeeId;         //Optional    Sales Employee
        public String SalesEmployeeName = "";
        public String ShipAgent = "";          //Optional N/A
        public String Location = "";           //Optional N/A
        public String BPType = "";             //Required    BP Type: C- Customer S- Vendor
        public String CreateTS = "";         //Required    Create Date & Time
        public String UpdateTS = "";           //Required    Update Date & Time

        public List<ShipToAddress> ShipToAddress = new List<ShipToAddress>();

        public Vendor(string code) : base()
        {
            _CardCode = code;

            string sql = String.Format(Src.Resource.Queries.GET_RECORD_OCRD, _CardCode.Replace("'", "''"));
            var rs = eCommon.ExecuteQuery(sql);
            if (rs.RecordCount == 0)
                throw new Exception($"Invalid Business Partner '{code}'");
            string cardType = rs.Fields.Item("CardType").Value.ToString().Trim().ToUpper();

            CardCode = rs.Fields.Item("CardCode").Value.ToString();
            CardName = rs.Fields.Item("CardName").Value.ToString();
            BPType = rs.Fields.Item("CardType").Value.ToString();
            if (BPType == "C")
                ObjType = "2C";
            else if (BPType == "S")
                ObjType = "2V";
            else
                throw new Exception("Invalid BP type. Currently only support Customer/Vendor");

            ForeignName = String.IsNullOrEmpty(rs.Fields.Item("CardFName").Value.ToString()) ? CardName : rs.Fields.Item("CardFName").Value.ToString();
            ShortName = rs.Fields.Item("CardName").Value.ToString();
            Telephone = rs.Fields.Item("Phone1").Value.ToString();

            //Change 1/0 required by Simplr
            //Active = rs.Fields.Item("validFor").Value.ToString() == "Y" || (rs.Fields.Item("validFor").Value.ToString() == "N" && rs.Fields.Item("frozenFor").Value.ToString() == "N") ? "Y" : "N";
            Active = rs.Fields.Item("validFor").Value.ToString() == "Y" || (rs.Fields.Item("validFor").Value.ToString() == "N" && rs.Fields.Item("frozenFor").Value.ToString() == "N") ? "1" : "0";

            FaxNumber = rs.Fields.Item("Fax").Value.ToString();
            Email = rs.Fields.Item("E_Mail").Value.ToString();
            WebSite = rs.Fields.Item("IntrntSite").Value.ToString();

            var CreateDate = (DateTime)rs.Fields.Item("CreateDate").Value;
            var UpdateDate = (DateTime)rs.Fields.Item("UpdateDate").Value;
            var CreateTime = (int)rs.Fields.Item("CreateTS").Value;
            var UpdateTime = (int)rs.Fields.Item("CreateTS").Value;

            CreateTS = eCommon.GetTimeStamp(CreateDate, CreateTime).ToString("yyyyMMdd HH:mm:ss");
            UpdateTS = eCommon.GetTimeStamp(UpdateDate, UpdateTime).ToString("yyyyMMdd HH:mm:ss");

            //Sales Employee
            sql = String.Format(Resource.Queries.GET_RECORD_OCRD_DEFAULT_SALES_EMPLOYEE, _CardCode.Replace("'", "''"));
            rs.DoQuery(sql);
            if (rs.Fields.Count > 0)
            {
                SalesEmployeeId = (int)rs.Fields.Item("SlpCode").Value;
                SalesEmployeeName = rs.Fields.Item("SlpName").Value.ToString();
            }

            sql = String.Format(Resource.Queries.GET_RECORD_OCRD_DEFAULT_CONTACT, _CardCode.Replace("'", "''"));
            rs.DoQuery(sql);
            if (rs.RecordCount > 0)
                ContactPerson = rs.Fields.Item("Name").Value.ToString();


            //BillTo
            sql = String.Format(Resource.Queries.GET_RECORD_CRD1_DEFAULT_BILL_TO, _CardCode.Replace("'", "''"));
            rs.DoQuery(sql);
            BillTo = String.IsNullOrEmpty(rs.Fields.Item("Address").Value.ToString()) ? "" : rs.Fields.Item("Address").Value.ToString();
            BillToStreet = String.IsNullOrEmpty(rs.Fields.Item("Street").Value.ToString()) ? "" : rs.Fields.Item("Street").Value.ToString();
            BillToBlock = String.IsNullOrEmpty(rs.Fields.Item("Block").Value.ToString()) ? "" : rs.Fields.Item("Block").Value.ToString();
            BillToPostCode = String.IsNullOrEmpty(rs.Fields.Item("ZipCode").Value.ToString()) ? "" : rs.Fields.Item("ZipCode").Value.ToString();
            BillToCity = String.IsNullOrEmpty(rs.Fields.Item("City").Value.ToString()) ? "" : rs.Fields.Item("City").Value.ToString();
            BillToCountry = String.IsNullOrEmpty(rs.Fields.Item("Country").Value.ToString()) ? "" : rs.Fields.Item("Country").Value.ToString();

            ShipToAddress = APIAccess.ShipToAddress.GetShipToAddress(CardCode);
        }

        internal override string sapKeyVal => this._CardCode.ToString();
        internal override string GetJsonObjectPayload() => Newtonsoft.Json.JsonConvert.SerializeObject(this);


    }
}
