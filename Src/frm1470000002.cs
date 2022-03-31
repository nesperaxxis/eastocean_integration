using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM;
using SAPbobsCOM;
using AXC_EOA_WMSIntegration.Src.APIAccess;

namespace AXC_EOA_WMSIntegration
{
    [Form("1470000002", false, "Bin Locations", "3072", 2)]
    [Authorization("1470000002_AXCSynch", "Synch Bin Locations", "", BoUPTOptions.bou_FullNone)]
    public class frm1470000002 : AXC_EOA_WMSIntegration.Src.Support.SystemForm
    {

        public frm1470000002() : base() {}
        public frm1470000002(String FormUID) : base(FormUID) {}
        public frm1470000002(SAPbouiCOM.Form oForm) : base(oForm) {}
        SAPbobsCOM.BoPermission isUserAuthorizeToSynch = BoPermission.boper_None;

        protected override void InitForm()
        {

        }


        protected override void GetItemReferences()
        {
            try
            {
                isUserAuthorizeToSynch = eCommon.GetUserAuthorization(SBOAddon.oCompany.UserSignature, "1470000002_AXCSynch");

                if (!SBOAddon.SBO_Application.Menus.Exists(SBOAddon.SYNCH_BIN_MENU_UID))
                    SBOAddon.SBO_Application.Menus.Item("1280").SubMenus.Add(SBOAddon.SYNCH_BIN_MENU_UID, SBOAddon.SYNCH_BIN_MENU_NAME, BoMenuType.mt_STRING, 99);

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

                _ = new frm1470000002(pVal.FormUID);
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
            frm1470000002 thisForm = SBOAddon.oOpenForms[pVal.FormUID] as frm1470000002;
            if (setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_BIN_LOCATIONS] != null && setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_BIN_LOCATIONS].ToString() == "Y")
            {
                //Auto synch is set to yes
                thisForm.Synch(true);
            }
        }

        [FormEvent("DataAddAfter", false)]
        public static void On_DataAddAfter(ref BusinessObjectInfo pVal)
        {
            axcFTSetup setup = new axcFTSetup();
            frm1470000002 thisForm = SBOAddon.oOpenForms[pVal.FormUID] as frm1470000002;
            if (setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_BIN_LOCATIONS] != null && setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_BIN_LOCATIONS].ToString() == "Y")
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
                if (_oForm.DataSources.DBDataSources.Item("OBIN").GetValue("BinCode", 0).ToString().Trim() == "")
                    throw new Exception("Invalid Bin code.");
                if (SBOAddon.SBO_Application.MessageBox("Synch this Bin code to WNS?", 2, "Ok", "Cancel") != 1)
                    throw new Exception("Synch operation canceled by user");

                SBOAddon.SBO_Application.StatusBar.SetText("Synching Bin code.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                _oForm.Freeze(true);

                bool isSuccess;
                string result;
                
                int thisItem = int.Parse(_oForm.DataSources.DBDataSources.Item("OBIN").GetValue("AbsEntry", 0).ToString().Trim());
                string thisBinCode = _oForm.DataSources.DBDataSources.Item("OBIN").GetValue("BinCode", 0).ToString().Trim();
                isSuccess = APIServiceAccess.SynchObject<Src.APIAccess.BinLocations>(thisItem, thisBinCode, out  result);
                if (isSuccess)
                    SBOAddon.SBO_Application.StatusBar.SetText("Bin successfully synched.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                else
                    SBOAddon.SBO_Application.StatusBar.SetText(String.Format("Posting failed. {0}", result), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            catch (Exception Ex)
            {
                SBOAddon.SBO_Application.StatusBar.SetText(Ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                _oForm.Freeze(false);
            }
        }

        
     }
}
