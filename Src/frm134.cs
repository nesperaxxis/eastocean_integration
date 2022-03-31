using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM;
using SAPbobsCOM;
using AXC_EOA_WMSIntegration.Src.APIAccess;

namespace AXC_EOA_WMSIntegration
{
    [Form("134", false, "Business Partners", "43535", 2)]
    [Authorization("134_AXCSynch", "WMS: Synch BP", "", BoUPTOptions.bou_FullNone)]
    public class frm134 : AXC_EOA_WMSIntegration.Src.Support.SystemForm
    {

        public frm134() : base() {}
        public frm134(String FormUID) : base(FormUID) {}
        public frm134(SAPbouiCOM.Form oForm) : base(oForm) {}

        protected override void InitForm()
        {
            
        }


        protected override void GetItemReferences()
        {
            try
            {
                if (!SBOAddon.SBO_Application.Menus.Exists(SBOAddon.SYNCH_BP_MENU_UID))
                    SBOAddon.SBO_Application.Menus.Item("1280").SubMenus.Add(SBOAddon.SYNCH_BP_MENU_UID, SBOAddon.SYNCH_BP_MENU_NAME, BoMenuType.mt_STRING, 99);

            }
            catch (Exception ex)
            { System.Windows.Forms.MessageBox.Show(ex.Message); }
        }

        protected override void OnBeforeFormClose(SBOItemEventArg pVal, out bool Bubble)
        {
            Bubble = true;
        }

        [FormEvent("LoadAfter", false)]
        public static void OnAfterFormLoad(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = SBOAddon.SBO_Application.Forms.Item(pVal.FormUID);
                //Add additional items on screen.

                _ = new frm134(pVal.FormUID);
            }
            catch (Exception ex)
            {
                SBOAddon.SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        [FormEvent("RightClickBefore", true)]
        public static void onRightClickBefore(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                //If menu to Synch to WMS does not exists, create it
                SAPbouiCOM.Form oForm = SBOAddon.SBO_Application.Forms.Item(eventInfo.FormUID);
                SBOAddon.SetMenuSynchEnable(oForm);
            }
            catch (Exception Ex)
            {
                SBOAddon.SBO_Application.StatusBar.SetText(Ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }

        }

        [FormEvent("DataUpdateAfter", false)]
        public static void On_DataUpdateAfter(ref BusinessObjectInfo pVal)
        {
            axcFTSetup setup = new axcFTSetup();
            frm134 thisForm = SBOAddon.oOpenForms[pVal.FormUID] as frm134;
            System.Diagnostics.Debug.WriteLine(thisForm._oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0).Trim());
            if (setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_BUSINESS_PARTNERS] != null && setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_BUSINESS_PARTNERS].ToString() == "Y"
                && thisForm._oForm.DataSources.DBDataSources.Item("OCRD").GetValue("U_AXC_SYNCH", 0).Trim() != "Y")
            {
                //Auto synch is set to yes and the synch to EP flag is set to 'Y' or Empty (default is to synch)
                thisForm.Synch(true);
            }
        }

        [FormEvent("DataAddAfter", false)]
        public static void On_DataAddAfter(ref BusinessObjectInfo pVal)
        {
            axcFTSetup setup = new axcFTSetup();
            frm134 thisForm = SBOAddon.oOpenForms[pVal.FormUID] as frm134;
            if (setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_BUSINESS_PARTNERS] != null && setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_BUSINESS_PARTNERS].ToString() == "Y"
                && thisForm._oForm.DataSources.DBDataSources.Item("OCRD").GetValue("U_AXC_SYNCH", 0).Trim() != "Y")
            {
                thisForm.Synch(true);
            }
        }

        /// <summary>
        /// Synch to WMS
        /// only in OK Mode
        /// only for Vendors
        /// </summary>
        public void Synch(bool isAuto=false)
        {
            try
            {
                bool isUserAuthorized = eCommon.GetExecuteAuthorizedEx("134_AXCSynch");
                if (!isUserAuthorized) return;

                //If auto post, by pass checking on the form mode.
                //Auto is always triggered after Add or Update.
                if (_oForm.DataSources.DBDataSources.Item("OCRD").GetValue("U_AXC_SYNCH", 0).ToString() == "Y")
                    throw new Exception("This BP is marked to exclude from integration.");
                if (!isAuto && _oForm.Mode != BoFormMode.fm_OK_MODE)
                    throw new Exception("Synch can only be done on 'OK' mode");
                if (_oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0).ToString().Trim() == "")
                    throw new Exception("Invalid vendor code.");
                if (SBOAddon.SBO_Application.MessageBox("Synch this vendor to WNS?", 2, "Ok", "Cancel") != 1)
                    throw new Exception("Synch operation canceled by user");

                SBOAddon.SBO_Application.StatusBar.SetText("Synching vendor.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                _oForm.Freeze(true);

                bool isSuccess;
                string result;
                
                string thisCardCode = _oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0).ToString().Trim();
                string thisCardType = _oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0).ToString().Trim();
                string thisCardName = _oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardName", 0).ToString().Trim();
                if (thisCardType != "C" && thisCardType != "S") throw new Exception("Only 'Customer'/'Vendor' is valid for updating to WMS");
                bool isVendor = thisCardType == "S";
                if (isVendor)
                    isSuccess = APIServiceAccess.SynchObject<Src.APIAccess.Vendor>(thisCardCode, thisCardName, out result);
                else
                    isSuccess = APIServiceAccess.SynchObject<Src.APIAccess.Customer>(thisCardCode, thisCardName, out result);
                if (isSuccess)
                    SBOAddon.SBO_Application.StatusBar.SetText($"Business Partner '{thisCardCode}' successfully synched.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                else 
                    throw new Exception(String.Format("Posting failed. {0}", result));
            }
            catch (Exception Ex)
            {
                SBOAddon.SBO_Application.StatusBar.SetText(Ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                if (isAuto)
                    SBOAddon.SBO_Application.MessageBox(Ex.Message);
            }
            finally
            {
                _oForm.Freeze(false);
            }
        }



        
     }
}
