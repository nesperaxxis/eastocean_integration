using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM;
using SAPbobsCOM;
using AXC_EOA_WMSIntegration.Src.APIAccess;

namespace AXC_EOA_WMSIntegration
{
    [Form("1474000001", false, "Inventory Counting", "3072", 2)]
    [Authorization("1474000001_AXCSynch", "Synch Inventory Counting", "", BoUPTOptions.bou_FullNone)]
    public class frm1474000001 : AXC_EOA_WMSIntegration.Src.Support.SystemForm
    {

        public frm1474000001() : base() {}
        public frm1474000001(String FormUID) : base(FormUID) {}
        public frm1474000001(SAPbouiCOM.Form oForm) : base(oForm) {}
        SAPbobsCOM.BoPermission isUserAuthorizeToSynch = BoPermission.boper_None;

        protected override void InitForm()
        {

        }

        protected override void GetItemReferences()
        {
            try
            {
                isUserAuthorizeToSynch = eCommon.GetUserAuthorization(SBOAddon.oCompany.UserSignature, "1474000001_AXCSynch");

                if (!SBOAddon.SBO_Application.Menus.Exists(SBOAddon.SYNCH_STOCK_COUNT_MENU_UID))
                    SBOAddon.SBO_Application.Menus.Item("1280").SubMenus.Add(SBOAddon.SYNCH_STOCK_COUNT_MENU_UID, SBOAddon.SYNCH_STOCK_COUNT_MENU_NAME, BoMenuType.mt_STRING, 99);

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

                _ = new frm1474000001(pVal.FormUID);
            }
            catch(Exception ex)
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
            frm1474000001 thisForm = SBOAddon.oOpenForms[pVal.FormUID] as frm1474000001;
            if (setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_STOCK_COUNTING] != null && setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_STOCK_COUNTING].ToString() == "Y")
            {
                //Auto synch is set to yes
                thisForm.Synch(true);
            }
        }

        [FormEvent("DataAddAfter", false)]
        public static void On_DataAddAfter(ref BusinessObjectInfo pVal)
        {
            axcFTSetup setup = new axcFTSetup();
            frm1474000001 thisForm = SBOAddon.oOpenForms[pVal.FormUID] as frm1474000001;
            if (setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_STOCK_COUNTING] != null && setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_STOCK_COUNTING].ToString() == "Y")
            {
                thisForm.Synch(true);
            }
        }

        /// <summary>
        /// Synch to WMS
        /// only in OK Mode
        /// </summary>
        public void Synch(bool isAuto=false)
        {
            try
            {
                if (isUserAuthorizeToSynch != BoPermission.boper_Full) return;              
                //If auto post, by pass checking on the form mode.
                //Auto is always triggered after Add or Update.
                if (!isAuto && _oForm.Mode != BoFormMode.fm_OK_MODE)
                    throw new Exception("Synch can only be done on 'OK' mode");
                if (_oForm.DataSources.DBDataSources.Item("OINC").GetValue("DocEntry", 0).ToString().Trim() == "")
                    throw new Exception("Invalid Inventory Counting doc entry.");
                if (_oForm.DataSources.DBDataSources.Item("OINC").GetValue("Status", 0).ToString().Trim() == "C")
                    throw new Exception("Inventory Counting status is closed/canceled.");
                if (SBOAddon.SBO_Application.MessageBox("Synch this Inventory Counting?", 2, "Ok", "Cancel") != 1)
                    throw new Exception("Synch operation canceled by user");

                SBOAddon.SBO_Application.StatusBar.SetText("Synching Inventory Counting.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                _oForm.Freeze(true);

                bool isSuccess;
                string result;
                
                int thisEntry =int.Parse( _oForm.DataSources.DBDataSources.Item("OINC").GetValue("DocEntry", 0).ToString().Trim());
                string thisNumber = _oForm.DataSources.DBDataSources.Item("OINC").GetValue("DocNum", 0).ToString().Trim();
                isSuccess = APIServiceAccess.SynchObject<Src.APIAccess.ItemCount>(thisEntry, thisNumber, out  result);
                if (isSuccess)
                    SBOAddon.SBO_Application.StatusBar.SetText("Inventory Counting successfully synched.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
