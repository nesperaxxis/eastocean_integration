using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM;
using SAPbobsCOM;
using AXC_EOA_WMSIntegration.Src.APIAccess;

namespace AXC_EOA_WMSIntegration
{
    [Form("63", false, "Item Groups", "43528", 2)]
    [Authorization("63_AXCSynch", "Synch Item Groups", "", BoUPTOptions.bou_FullNone)]
    public class frm63 : AXC_EOA_WMSIntegration.Src.Support.SystemForm
    {

        public frm63() : base() {}
        public frm63(String FormUID) : base(FormUID) {}
        public frm63(SAPbouiCOM.Form oForm) : base(oForm) {}
        public int removingCode = 0;
        SAPbobsCOM.BoPermission isUserAuthorizeToSynch = BoPermission.boper_None;

        protected override void InitForm()
        {
        }


        protected override void GetItemReferences()
        {
            try
            {
                isUserAuthorizeToSynch = eCommon.GetUserAuthorization(SBOAddon.oCompany.UserSignature, "63_AXCSynch");
                if (!SBOAddon.SBO_Application.Menus.Exists(SBOAddon.SYNCH_ITEM_CATEGORY_MENU_UID))
                    SBOAddon.SBO_Application.Menus.Item("1280").SubMenus.Add(SBOAddon.SYNCH_ITEM_CATEGORY_MENU_UID, SBOAddon.SYNCH_ITEM_CATEGORY_MENU_NAME, BoMenuType.mt_STRING, 99);

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

                _ = new frm63(pVal.FormUID);
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
            frm63 thisForm = SBOAddon.oOpenForms[pVal.FormUID] as frm63;
            if (setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_ITEM_GROUPS] != null && setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_ITEM_GROUPS].ToString() == "Y")
            {
                //Auto synch is set to yes
                thisForm?.Synch(true);
            }
        }

        [FormEvent("DataAddAfter", false)]
        public static void On_DataAddAfter(ref BusinessObjectInfo pVal)
        {
            axcFTSetup setup = new axcFTSetup();
            frm63 thisForm = SBOAddon.oOpenForms[pVal.FormUID] as frm63;
            if (setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_ITEM_GROUPS] != null && setup.Values[SBOAddon_DB.OFTIS_WS_AUTO_SYNCH_ITEM_GROUPS].ToString() == "Y")
            {
                thisForm?.Synch(true);
            }
        }

        /// <summary>
        /// Synch to WMS
        /// only in OK Mode
        /// </summary>
        public void Synch(bool isAuto = false)
        {
            try
            {
                if (isUserAuthorizeToSynch != BoPermission.boper_Full) return;
                //If auto post, by pass checking on the form mode.
                //Auto is always triggered after Add or Update.
                int thisGroup = Convert.ToInt32( _oForm.DataSources.DBDataSources.Item("OITB").GetValue("ItmsGrpCod", 0).ToString().Trim());

                if (!isAuto && _oForm.Mode != BoFormMode.fm_OK_MODE)
                    throw new Exception("Synch can only be done on 'OK' mode");
                if (thisGroup == 0)
                    throw new Exception("Invalid Item Group.");
                if (SBOAddon.SBO_Application.MessageBox("Synch this Item Group to WNS?", 2, "Ok", "Cancel") != 1)
                    throw new Exception("Synch operation canceled by user");

                SBOAddon.SBO_Application.StatusBar.SetText("Synching Item Groups.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                _oForm.Freeze(true);

                bool isSuccess;
                string result;
                string thisGroupName = _oForm.DataSources.DBDataSources.Item("OITB").GetValue("ItmsGrpNam", 0).ToString().Trim();
                isSuccess = APIServiceAccess.SynchObject<Src.APIAccess.ItemCategories>(thisGroup, thisGroupName, out result);
                if (isSuccess)
                    SBOAddon.SBO_Application.StatusBar.SetText("Item Group successfully synched.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
