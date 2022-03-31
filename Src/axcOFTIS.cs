using System;
using System.Collections.Generic;
using System.Text;
using AXC_EOA_WMSIntegration.Src.APIAccess;
using SAPbouiCOM;

namespace AXC_EOA_WMSIntegration
{
    [Form("axcOFTIS", true, "Setup", "", 1)]
    [Authorization("axcOFTIS", "Setup", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)]
    public class axcOFTIS : AXC_EOA_WMSIntegration.Src.Support.UserForm
    {
        //SAPbouiCOM.Form _oForm = null;

        SAPbouiCOM.Button _btnOK = null;
        SAPbouiCOM.DBDataSource _dbOFTIS = null;
        SAPbouiCOM.DBDataSource _dbFTIS1 = null;
        SAPbouiCOM.Matrix _mtxFTIS1 = null;
        
        public axcOFTIS() : base()
        {
        }

        public axcOFTIS(SAPbouiCOM.Form oForm) : base (oForm)
        {
        }

        public axcOFTIS(String FormUID) : base (FormUID)
        {
        }

        protected override void InitForm()
        {
            _oForm.Freeze(true);
            try
            {
                //Form only for ok mode
                _oForm.SupportedModes = 9;
                //Check if the setting record exists, if not add it.
                axcFTSetup ftSetup = new axcFTSetup();
                //Call the record on this screen.
                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oCon = oCons.Add();
                oCon.Alias = "Code";
                oCon.Operation = BoConditionOperation.co_EQUAL;
                oCon.CondVal = SBOAddon_DB.SETUP_UDO_RECORD_ID;

                _dbOFTIS.Query(oCons);
                _dbFTIS1.Query(oCons);
                _mtxFTIS1.LoadFromDataSource();

                _oForm.Items.Item("tabTM_WS").Click(BoCellClickType.ct_Regular);
                
            }
            finally
            {
                _oForm.Freeze(false);
            }
        }

        protected override void GetItemReferences()
        {
            _mtxFTIS1 = _oForm.Items.Item("mtxFTIS1").Specific as SAPbouiCOM.Matrix;
            _dbOFTIS = _oForm.DataSources.DBDataSources.Item("@AXC_OFTIS");
            _dbFTIS1 = _oForm.DataSources.DBDataSources.Item("@AXC_FTIS1");

            _btnOK = _oForm.Items.Item("1").Specific as SAPbouiCOM.Button;

            //The Events
            _btnOK.PressedBefore += _btnOK_PressedBefore;
            _btnOK.PressedAfter += _btnOK_PressedAfter;
        }

        void _btnOK_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            try
            {
                BubbleEvent = true;
            }
            catch (Exception Exc)
            {
                BubbleEvent = false;
                AXC_EOA_WMSIntegration.Src.Support.Addon.SBO_Application.StatusBar.SetText(Exc.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        void _btnOK_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                APIServiceAccess.newWsService();
            }
            catch(Exception Exc)
            {
                AXC_EOA_WMSIntegration.Src.Support.Addon.SBO_Application.StatusBar.SetText(Exc.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

  
        [FormEvent("ResizeAfter",false)]
        public static void OnAfterFormResize(SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = AXC_EOA_WMSIntegration.Src.Support.Addon.SBO_Application.Forms.Item(pVal.FormUID);
            if (oForm.Items.Count > 0)
            {
                //oForm.Items.Item("Item_4").Width = oForm.ClientWidth - 10;
                //oForm.Items.Item("Item_4").Height = oForm.Items.Item("1").Top - 10 - oForm.Items.Item("Item_4").Top;
            }
        }

        protected override void OnBeforeFormClose(SBOItemEventArg pVal, out bool Bubble)
        {
            Bubble = true;
            //Remove all the events
            _btnOK.PressedBefore -= _btnOK_PressedBefore;
            
        }


    }

    public class axcFTSetup
    {

        public System.Collections.Generic.Dictionary<string, object> Values = new Dictionary<string, object>();
        private int unMappedUser = 0;

        public axcFTSetup()
        {
            LoadRecord();
            if (Values.Count == 0)
                GenerateRecord(true);
            else if (unMappedUser > 0)
                GenerateRecord(false);
            
        }

        private void LoadRecord()
        {
            SAPbobsCOM.Recordset oRS = SBOAddon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            oRS.DoQuery(String.Format("SELECT * FROM [@AXC_OFTIS] WHERE Code = '{0}'", SBOAddon_DB.SETUP_UDO_RECORD_ID));
            if (oRS.RecordCount > 0)
            {
                Values = new Dictionary<string, object>();
                for (int iField = 0; iField < oRS.Fields.Count; iField++)
                {
                    Values.Add(oRS.Fields.Item(iField).Name, oRS.Fields.Item(iField).Value);                    
                }
            }

            oRS.DoQuery(Src.Resource.Queries.FTIS1_GET_UNMAPPED_USER);
            unMappedUser = oRS.RecordCount;

            eCommon.ReleaseComObject(oRS);
        }

        public void GenerateRecord(bool isNew = true)
        {
            SAPbobsCOM.CompanyService oCS = SBOAddon.oCompany.GetCompanyService();
            SAPbobsCOM.GeneralService oGS = oCS.GetGeneralService(SBOAddon_DB.UDO_CODE_SETUP);
            SAPbobsCOM.GeneralData oGD = oGS.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData) as SAPbobsCOM.GeneralData;
            SAPbobsCOM.GeneralDataCollection oGDC = null;
            SAPbobsCOM.GeneralData oGD1 = null;
            SAPbobsCOM.GeneralDataParams oGP = oGS.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams) as SAPbobsCOM.GeneralDataParams;
            SAPbobsCOM.Recordset oRS = SBOAddon.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;

            if (isNew)
            {
                oGD = oGS.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData) as SAPbobsCOM.GeneralData;
                oGD.SetProperty("Code", SBOAddon_DB.SETUP_UDO_RECORD_ID);
                oGD.SetProperty("Name", SBOAddon_DB.SETUP_UDO_RECORD_ID);
            } else
            {
                oGP.SetProperty("Code", SBOAddon_DB.SETUP_UDO_RECORD_ID);
                oGD = oGS.GetByParams(oGP);
            }
            bool isUpdated = false;
            foreach (String key in Values.Keys)
            {
                if (!key.StartsWith("U_"))
                    continue;
                if (Values[key] == null)
                    continue;
                
                if (!oGD.GetProperty(key).Equals(Values[key])) { oGD.SetProperty(key, Values[key]); isUpdated = true; }
            }

            String sSQL = Src.Resource.Queries.FTIS1_GET_UNMAPPED_USER;
            oRS.DoQuery(sSQL);
            oGDC = oGD.Child("AXC_FTIS1");
            for (int i =0; i < oRS.RecordCount; i++)
            {
                oGD1 = oGDC.Add();
                oGD1.SetProperty(SBOAddon_DB.FTIS1_ALERT_USERID, oRS.Fields.Item("USERID").Value.ToString());
                oGD1.SetProperty(SBOAddon_DB.FTIS1_ALERT_USERNAME, oRS.Fields.Item("U_NAME").Value.ToString());

                isUpdated = true;
                oRS.MoveNext();
            }


            if (isNew)
                oGP = oGS.Add(oGD);
            else if (isUpdated)
                oGS.Update(oGD);

            eCommon.ReleaseComObject(oGD);
            eCommon.ReleaseComObject(oGS);
            eCommon.ReleaseComObject(oCS);
            eCommon.ReleaseComObject(oRS);

            LoadRecord();
        }

    }
}
