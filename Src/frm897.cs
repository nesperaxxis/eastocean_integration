using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM;
using SAPbobsCOM;
using AXC_EOA_WMSIntegration.Src.APIAccess;

namespace AXC_EOA_WMSIntegration
{
    [Form("897", false, "Manufacturers", "11520", 99)]
    [Authorization("897_AXCSynch", "Synch Brands", "", BoUPTOptions.bou_FullNone)]
    public class frm897 : AXC_EOA_WMSIntegration.Src.Support.SystemForm
    {
        private SAPbouiCOM.Matrix _mtx3 = null;
        private SAPbouiCOM.UserDataSource _udsClickedLine = null;

        public frm897() : base() {}
        public frm897(String FormUID) : base(FormUID) {}
        public frm897(SAPbouiCOM.Form oForm) : base(oForm) { }
        SAPbobsCOM.BoPermission isUserAuthorizeToSynch = BoPermission.boper_None;

        protected override void InitForm()
        {

        }


        protected override void GetItemReferences()
        {
            try
            {
                try
                {
                    isUserAuthorizeToSynch = eCommon.GetUserAuthorization(SBOAddon.oCompany.UserSignature, "897_AXCSynch");
                    _udsClickedLine = _oForm.DataSources.UserDataSources.Item("axcClicked");
                }
                catch
                {
                    _udsClickedLine = _oForm.DataSources.UserDataSources.Add("axcClicked", BoDataType.dt_SHORT_NUMBER);
                }

                if (!SBOAddon.SBO_Application.Menus.Exists(SBOAddon.SYNCH_BRAND_MENU_UID))
                    SBOAddon.SBO_Application.Menus.Item("1280").SubMenus.Add(SBOAddon.SYNCH_BRAND_MENU_UID, SBOAddon.SYNCH_BRAND_MENU_NAME, BoMenuType.mt_STRING, 99);

                _mtx3 = _oForm.Items.Item("3").Specific as SAPbouiCOM.Matrix;

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

                _ = new frm897(pVal.FormUID);
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
                oForm.DataSources.UserDataSources.Item("axcClicked").ValueEx = eventInfo.Row.ToString();
            }
            catch (Exception Ex)
            {
                SBOAddon.SBO_Application.StatusBar.SetText(Ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }

        }


        /// <summary>
        /// Synch to WMS
        /// only in OK Mode
        /// </summary>
        public void Synch()
        {
            try
            {
                if (isUserAuthorizeToSynch != BoPermission.boper_Full) return;

                if (_oForm.Mode != BoFormMode.fm_OK_MODE)
                    throw new Exception("Synch can only be done on 'OK' mode");


                int iSelectedRow = 0;
                if(!int.TryParse(_udsClickedLine.ValueEx, out iSelectedRow))
                    throw new Exception("Please select a manufacturer to synch.");

                string thisBrand = (_mtx3.GetCellSpecific("FirmName", iSelectedRow) as SAPbouiCOM.EditText).String;
                int thisBrandId = int.Parse((_mtx3.GetCellSpecific("FirmCode", iSelectedRow) as SAPbouiCOM.EditText).String);
                if (thisBrand.Trim() == "")
                    return;


                if (SBOAddon.SBO_Application.MessageBox(String.Format("Synch Brand '{0}' to WNS?", thisBrand), 2, "Ok", "Cancel") != 1)
                    throw new Exception("Synch operation canceled by user");

                SBOAddon.SBO_Application.StatusBar.SetText(String.Format("Synching brand '{0}'.", thisBrand), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                _oForm.Freeze(true);

                bool isSuccess;
                string result;
                isSuccess = APIServiceAccess.SynchObject<Brands>(thisBrandId, thisBrand, out result);
                if (isSuccess)
                    SBOAddon.SBO_Application.StatusBar.SetText("Brand successfully synched.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
