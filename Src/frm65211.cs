using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM;
using SAPbobsCOM;
using AXC_EOA_WMSIntegration.Src.APIAccess;

namespace AXC_EOA_WMSIntegration
{
    [Form("65211", false, "Work Order", "3072", 2)]
    [Authorization("65211_AXCSynch", "Synch Work Orders", "", BoUPTOptions.bou_FullNone)]
    public class frm65211 : AXC_EOA_WMSIntegration.Src.Support.SystemForm
    {

        public frm65211() : base() {}
        public frm65211(String FormUID) : base(FormUID) {}
        public frm65211(SAPbouiCOM.Form oForm) : base(oForm) {}
        SAPbobsCOM.BoPermission isUserAuthorizeToSynch = BoPermission.boper_None;

        protected override void InitForm()
        {

        }

        protected override void GetItemReferences()
        {
            try
            {
                isUserAuthorizeToSynch = eCommon.GetUserAuthorization(SBOAddon.oCompany.UserSignature, "65211_AXCSynch");

                if (!SBOAddon.SBO_Application.Menus.Exists(SBOAddon.SYNCH_WORK_ORDER_MENU_UID))
                    SBOAddon.SBO_Application.Menus.Item("1280").SubMenus.Add(SBOAddon.SYNCH_WORK_ORDER_MENU_UID, SBOAddon.SYNCH_WORK_ORDER_MENU_NAME, BoMenuType.mt_STRING, 99);

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

                _ = new frm65211(pVal.FormUID);
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
            frm65211 thisForm = SBOAddon.oOpenForms[pVal.FormUID] as frm65211;
            if (setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_WORK_ORDERS] != null && setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_WORK_ORDERS].ToString() == "Y")
            {
                //Auto synch is set to yes
                thisForm?.Synch(true);
            }
        }

        [FormEvent("DataAddAfter", false)]
        public static void On_DataAddAfter(ref BusinessObjectInfo pVal)
        {
            axcFTSetup setup = new axcFTSetup();
            frm65211 thisForm = SBOAddon.oOpenForms[pVal.FormUID] as frm65211;
            if (setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_WORK_ORDERS] != null && setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_WORK_ORDERS].ToString() == "Y")
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

                //If auto post, by pass checking on the form mode.
                //Auto is always triggered after Add or Update.
                if (!isAuto && _oForm.Mode != BoFormMode.fm_OK_MODE)
                    throw new Exception("Synch can only be done on 'OK' mode");
                if (_oForm.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0).ToString().Trim() == "")
                    throw new Exception("Invalid Production Order doc entry.");
                if (_oForm.DataSources.DBDataSources.Item("OWOR").GetValue("Status", 0).ToString().Trim() != "R")
                    throw new Exception("Document status not 'Released'.");
                if (SBOAddon.SBO_Application.MessageBox("Synch this production order?", 2, "Ok", "Cancel") != 1)
                    throw new Exception("Synch operation canceled by user");

                SBOAddon.SBO_Application.StatusBar.SetText("Synching production order.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                _oForm.Freeze(true);

                bool isSuccess;
                string result;
                
                int thisEntry =int.Parse( _oForm.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0).ToString().Trim());
                string thisNumber = _oForm.DataSources.DBDataSources.Item("OWOR").GetValue("DocNum", 0).ToString().Trim();
                isSuccess = APIServiceAccess.SynchObject<Src.APIAccess.WorkOrders>(thisEntry, thisNumber, out  result);
                if (isSuccess)
                    SBOAddon.SBO_Application.StatusBar.SetText("Production Order successfully synched.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                else
                {
                    SBOAddon.SBO_Application.StatusBar.SetText(String.Format("Posting failed. {0}", result), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    if (isAuto)
                        SBOAddon.SBO_Application.MessageBox(String.Format("Posting failed. {0}", result));
                }
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
