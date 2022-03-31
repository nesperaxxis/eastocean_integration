using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM;
using SAPbobsCOM;
using AXC_EOA_WMSIntegration.Src.APIAccess;

namespace AXC_EOA_WMSIntegration
{
    [Form("180", false, "AR Return", "3072", 2)]
    [Authorization("180_AXCSynch", "Synch AR Return", "", BoUPTOptions.bou_FullNone)]
    public class frm180 : AXC_EOA_WMSIntegration.Src.Support.SystemForm
    {

        public frm180() : base() {}
        public frm180(String FormUID) : base(FormUID) {}
        public frm180(SAPbouiCOM.Form oForm) : base(oForm) {}
        SAPbobsCOM.BoPermission isUserAuthorizeToSynch = BoPermission.boper_None;

        protected override void InitForm()
        {

        }


        protected override void GetItemReferences()
        {
            try
            {
                isUserAuthorizeToSynch = eCommon.GetUserAuthorization(SBOAddon.oCompany.UserSignature, "180_AXCSynch");

                if (!SBOAddon.SBO_Application.Menus.Exists(SBOAddon.SYNCH_AR_RETURN_MENU_UID))
                    SBOAddon.SBO_Application.Menus.Item("1280").SubMenus.Add(SBOAddon.SYNCH_AR_RETURN_MENU_UID, SBOAddon.SYNCH_AR_RETURN_MENU_NAME, BoMenuType.mt_STRING, 99);

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

                _ = new frm180(pVal.FormUID);
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
            frm180 thisForm = SBOAddon.oOpenForms[pVal.FormUID] as frm180;
            if (setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_AR_RETURNS] != null && setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_AR_RETURNS].ToString() == "Y")
            {
                //Auto synch is set to yes
                thisForm?.Synch(true);
            }
        }

        [FormEvent("DataAddAfter", false)]
        public static void On_DataAddAfter(ref BusinessObjectInfo pVal)
        {
            axcFTSetup setup = new axcFTSetup();
            frm180 thisForm = SBOAddon.oOpenForms[pVal.FormUID] as frm180;
            if (setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_AR_RETURNS] != null && setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_AR_RETURNS].ToString() == "Y")
            {
                thisForm?.Synch(true);
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
                if (isAuto && _oForm.DataSources.DBDataSources.Item("ORDN").GetValue("DocType", 0).ToString().Trim() != "I")
                    return;
                else if (!isAuto && _oForm.DataSources.DBDataSources.Item("ORDN").GetValue("DocType", 0).ToString().Trim() != "I")
                    throw new Exception("Only DocType = 'Item' can be synch.");


                //If auto post, by pass checking on the form mode.
                //Auto is always triggered after Add or Update.
                if (!isAuto && _oForm.Mode != BoFormMode.fm_OK_MODE)
                    throw new Exception("Synch can only be done on 'OK' mode");
                if (_oForm.DataSources.DBDataSources.Item("ORDN").GetValue("DocEntry", 0).ToString().Trim() == "")
                    throw new Exception("Invalid Return doc entry.");
                if (_oForm.DataSources.DBDataSources.Item("ORDN").GetValue(SBOAddon_DB.ODOC_UDF_EXTERNAL_ID, 0).ToString().Trim() != "")
                    throw new Exception("Return Document was originated from WMS, cannot push back to WMS.");
                if (_oForm.DataSources.DBDataSources.Item("ORDN").GetValue("CANCELED", 0).ToString().Trim() != "N")
                    throw new Exception("Document is canceled/cancellation document.");
                if (SBOAddon.SBO_Application.MessageBox("Synch this Return?", 2, "Ok", "Cancel") != 1)
                    throw new Exception("Synch operation canceled by user");

                SBOAddon.SBO_Application.StatusBar.SetText("Synching Return.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                _oForm.Freeze(true);

                bool isSuccess;
                string result;

                string thisNumber = _oForm.DataSources.DBDataSources.Item("ORDN").GetValue("DocNum", 0).ToString().Trim();
                int thisEntry =int.Parse( _oForm.DataSources.DBDataSources.Item("ORDN").GetValue("DocEntry", 0).ToString().Trim());
                isSuccess = APIServiceAccess.SynchObject<Src.APIAccess.ARReturns>(thisEntry, thisNumber, out  result);
                if (isSuccess)
                    SBOAddon.SBO_Application.StatusBar.SetText("Return successfully synched.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
