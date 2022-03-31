using System;
using System.Collections.Generic;
using System.Windows.Forms;
using SAPbouiCOM;
using SAPbobsCOM;

/*
 * This project depends on 
 * NUGET Microsoft.AspNet.WebApi.Client 2.2
 * 
 * Version History :
 *  
*/
namespace AXC_EOA_WMSIntegration
{
    class SBOAddon : Src.Support.Addon
    {
        public const string SYNCH_O_OBJECT_VENDOR = "2V";
        public const string SYNCH_O_OBJECT_CUSTOMER = "2C";
        public const string SYNCH_O_OBJECT_ITEM = "4";
        public const string SYNCH_O_OBJECT_ITEM_CATEGORY = "52";
        public const string SYNCH_O_OBJECT_BRAND = "43";
        public const string SYNCH_O_OBJECT_BAR_CODE = "1470000062";
        public const string SYNCH_O_OBJECT_WAREHOUSE = "64";
        public const string SYNCH_O_OBJECT_BIN = "10000206";
        public const string SYNCH_O_OBJECT_BOM = "66";
        public const string SYNCH_O_OBJECT_SALES_ORDER = "17";
        public const string SYNCH_O_OBJECT_RESERVE_INVOICE = "13R";
        public const string SYNCH_O_OBJECT_AR_CN = "14";
        public const string SYNCH_O_OBJECT_AR_RETURNS = "16";
        public const string SYNCH_O_OBJECT_PURCHASE_ORDER = "22";
        public const string SYNCH_O_OBJECT_AP_RETURN = "21";
        public const string SYNCH_O_OBJECT_AP_CN = "19";
        public const string SYNCH_O_OBJECT_WORK_ORDER = "202";
        public const string SYNCH_I_OBJECT_PICK_LIST = "156";
        public const string SYNCH_I_OBJECT_AR_RETURN = "16";
        public const string SYNCH_I_OBJECT_GRPO = "20";
        public const string SYNCH_I_OBJECT_TR_REQUEEST = "1250000001";
        public const string SYNCH_I_OBJECT_WHS_TRANSFER = "67";
        public const string SYNCH_I_OBJECT_ISSUE_PROD = "60P";
        public const string SYNCH_I_OBJECT_RECPT_PROD = "59P";
        public const string SYNCH_I_OBJECT_STOCK_ADJ_NEG = "60";
        public const string SYNCH_I_OBJECT_STOCK_ADJ_POS = "59";
        public const string SYNCH_O_OBJECT_STOCK_COUNT = "1470000065";
        public const string SYNCH_I_OBJECT_STOCK_COUNT = "10000071";


        public const string SYNCH_BP_MENU_UID = "AXC_SYNCH_CRD";
        public const string SYNCH_BP_MENU_NAME = "WS: Synch Vendor";
        public const string SYNCH_ITEM_MENU_UID = "AXC_SYNCH_ITM";
        public const string SYNCH_ITEM_MENU_NAME = "WS: Synch Item";
        public const string SYNCH_ITEM_CATEGORY_MENU_UID = "AXC_SYNCH_ITB";
        public const string SYNCH_ITEM_CATEGORY_MENU_NAME = "WS: Synch Item Categories";
        public const string SYNCH_ITEM_BARCODES_MENU_UID = "AXC_SYNCH_BCD";
        public const string SYNCH_ITEM_BARCODES_MENU_NAME = "WS: Synch Item Barcodes";
        public const string SYNCH_WAREHOUSE_MENU_UID = "AXC_SYNCH_WHS";
        public const string SYNCH_WAREHOUSE_MENU_NAME = "WS: Synch Warehouse";
        public const string SYNCH_BIN_MENU_UID = "AXC_SYNCH_BIN";
        public const string SYNCH_BIN_MENU_NAME = "WS: Synch Bin Location";
        public const string SYNCH_BOM_MENU_UID = "AXC_SYNCH_ITT";
        public const string SYNCH_BOM_MENU_NAME = "WS: Synch BOM";
        public const string SYNCH_BRAND_MENU_UID = "AXC_SYNCH_MRC";
        public const string SYNCH_BRAND_MENU_NAME = "WS: Synch Brand";
        public const string SYNCH_SALES_ORDER_MENU_UID = "AXC_SYNCH_RDR";
        public const string SYNCH_SALES_ORDER_MENU_NAME = "WS: Synch SO";
        public const string SYNCH_RES_INVOICE_MENU_UID = "AXC_SYNCH_INV";
        public const string SYNCH_RES_INVOICE_MENU_NAME = "WS: Synch Rsrv Invoice";
        public const string SYNCH_AR_CN_MENU_UID = "AXC_SYNCH_RIN";
        public const string SYNCH_AR_CN_MENU_NAME = "WS: Synch AR CN";
        public const string SYNCH_AR_RETURN_MENU_UID = "AXC_SYNCH_RDN";
        public const string SYNCH_AR_RETURN_MENU_NAME = "WS: Synch Return";
        public const string SYNCH_PO_MENU_UID = "AXC_SYNCH_POR";
        public const string SYNCH_PO_MENU_NAME = "WS: Synch PO";
        public const string SYNCH_AP_RETURN_MENU_UID = "AXC_SYNCH_RPD";
        public const string SYNCH_AP_RETURN_MENU_NAME = "WS: Synch Return";
        public const string SYNCH_AP_CN_MENU_UID = "AXC_SYNCH_RPC";
        public const string SYNCH_AP_CN_MENU_NAME = "WS: Synch AP CN";
        public const string SYNCH_WORK_ORDER_MENU_UID = "AXC_SYNCH_WOR";
        public const string SYNCH_WORK_ORDER_MENU_NAME = "WS: Synch Production Order";
        public const string SYNCH_STOCK_COUNT_MENU_UID = "AXC_SYNCH_INC";
        public const string SYNCH_STOCK_COUNT_MENU_NAME = "WS: Synch Item Count";




        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(String[] Args)
        {
            try
            {
                var addon = new SBOAddon(Args, "AXC_EOA_WMSIntegration", "WMS Integration");
                if (addon.Connected)
                    System.Windows.Forms.Application.Run();

            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message);
            }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        public SBOAddon(String[] Args, String AddonCode, String AddonName) : base(Args, AddonCode, AddonName)
        {
            try
            {
                SBO_Application.RightClickEvent += SBO_Application_RightClickEvent;


                //Notify the users the addon is ready to use.
                SBO_Application.StatusBar.SetText("Addon " + SBOAddon.gcAddOnName + " is ready.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                Connected = true;
            }
            catch (Exception ex)
            {
                Connected = false;
                MessageBox.Show("Failed initializing addon. " + ex.Message);
            }
            finally
            {
            }

        }

        protected override void CreateMenuTree()
        {
            base.CreateMenuTree();
            

        }

        private void SBO_Application_RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (eventInfo.BeforeAction)
                {
                    SAPbouiCOM.Form oForm = SBOAddon.SBO_Application.Forms.Item(eventInfo.FormUID);
                    SBOAddon.SetMenuSynchEnable(oForm);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }


        protected override void OnAppEvents(BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case BoAppEventTypes.aet_CompanyChanged:
                case BoAppEventTypes.aet_ServerTerminition:
                case BoAppEventTypes.aet_ShutDown:
                    try
                    {
                        if (SBO_Application.Menus.Exists(SBOAddon.gcAddOnName)) SBO_Application.Menus.RemoveEx(SBOAddon.gcAddOnName);
                    }
                    catch { }
                    System.Windows.Forms.Application.Exit();
                    break;
                case BoAppEventTypes.aet_FontChanged:
                    break;
                case BoAppEventTypes.aet_LanguageChanged:
                    break;

            }


        }

        protected override void OnMenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool Bubble)
        {
            Bubble = true;
            try
            {
                if (pVal.BeforeAction == true)
                {
                    SAPbouiCOM.Form oForm = null;

                    oForm = SBO_Application.Forms.ActiveForm;
                    String sXML = oForm.GetAsXML();
                    switch (pVal.MenuUID)
                    {
                        case "1283":    //Delete Record
                            break;
                        case "1293":
                            break;
                        case "1285":        //Restore Form
                            break;
                    }
                }
                else
                {
                    //After Menu
                    SAPbouiCOM.Form oActiveForm = null;
                    switch (pVal.MenuUID)
                    {
                        case "1293":        //Delete Row Menu
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            switch (oActiveForm.TypeEx)
                            {
                                default:
                                    break;
                            }
                            break;
                        case "1283":
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            switch (oActiveForm.TypeEx)
                            {
                                default:
                                    break;
                            }
                            break;

                        case "1282":    //Add Menu pressed
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            switch (oActiveForm.TypeEx)
                            {
                                default:
                                    break;
                            }
                            break;
                        case "1281":   //Find Menu
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            switch (oActiveForm.TypeEx)
                            {
                                default:
                                    break;
                            }
                            break;
                        case "1280":    //DataMenu
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            break;
                        case SBOAddon.SYNCH_BP_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm134 _oForm134 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm134;
                            if (_oForm134 != null)
                                _oForm134.Synch();
                            break;
                        case SBOAddon.SYNCH_ITEM_BARCODES_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm1470000020 _oForm1470000020 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm1470000020;
                            if (_oForm1470000020 != null)
                                _oForm1470000020.Synch();
                            break;
                        case SBOAddon.SYNCH_ITEM_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm150 _oForm150 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm150;
                            if (_oForm150 != null)
                                _oForm150.Synch();
                            break;
                        case SBOAddon.SYNCH_ITEM_CATEGORY_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm63 _oForm63 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm63;
                            if (_oForm63 != null)
                                _oForm63.Synch();
                            break;
                        case SBOAddon.SYNCH_BRAND_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm897 _oForm897 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm897;
                            if (_oForm897 != null)
                                _oForm897.Synch();
                            break;
                        case SBOAddon.SYNCH_WAREHOUSE_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm62 _oForm62 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm62;
                            if (_oForm62 != null)
                                _oForm62.Synch();
                            break;
                        case SBOAddon.SYNCH_BIN_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm1470000002 _frm1470000002 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm1470000002;
                            if (_frm1470000002 != null)
                                _frm1470000002.Synch();
                            break;
                        case SBOAddon.SYNCH_BOM_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm672 _frm672 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm672;
                            if (_frm672 != null)
                                _frm672.Synch();
                            break;
                        case SBOAddon.SYNCH_WORK_ORDER_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm65211 _frm65211 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm65211;
                            if (_frm65211 != null)
                                _frm65211.Synch();
                            break;
                        case SBOAddon.SYNCH_SALES_ORDER_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm139 _frm139 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm139;
                            if (_frm139 != null)
                                _frm139.Synch();
                            break;
                        case SBOAddon.SYNCH_RES_INVOICE_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm60091 _frm60091 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm60091;
                            if (_frm60091 != null)
                                _frm60091.Synch();
                            break;
                        case SBOAddon.SYNCH_AR_RETURN_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm180 _frm180 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm180;
                            if (_frm180 != null)
                                _frm180.Synch();
                            break;
                        case SBOAddon.SYNCH_AR_CN_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm179 _frm179 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm179;
                            if (_frm179 != null)
                                _frm179.Synch();
                            break;
                        case SBOAddon.SYNCH_PO_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm142 _frm142 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm142;
                            if (_frm142 != null)
                                _frm142.Synch();
                            break;
                        case SBOAddon.SYNCH_AP_RETURN_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm182 _frm182 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm182;
                            if (_frm182 != null)
                                _frm182.Synch();
                            break;
                        case SBOAddon.SYNCH_AP_CN_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm181 _frm181 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm181;
                            if (_frm181 != null)
                                _frm181.Synch();
                            break;
                        case SBOAddon.SYNCH_STOCK_COUNT_MENU_UID:
                            oActiveForm = SBO_Application.Forms.ActiveForm;
                            frm1474000001 _oForm1474000001 = SBOAddon.oOpenForms[oActiveForm.UniqueID] as frm1474000001;
                            if (_oForm1474000001 != null)
                                _oForm1474000001.Synch();
                            break;
                        default:
                            FormAttribute oAttrib = Forms[pVal.MenuUID] as FormAttribute;
                            if (oAttrib != null)
                            {
                                try
                                {
                                    //Execute the constructor
                                    System.Reflection.Assembly asm = System.Reflection.Assembly.GetExecutingAssembly();
                                    Type oType = asm.GetType(oAttrib.TypeName);
                                    System.Reflection.ConstructorInfo ctor = oType.GetConstructor(new Type[0]);
                                    if (ctor != null)
                                    {
                                        object oForm = ctor.Invoke(new Object[0]);
                                    }
                                    else
                                        throw new Exception("No default constructor found for form type - " + oAttrib.FormType);
                                }
                                catch (Exception Ex)
                                {
                                    SBO_Application.MessageBox(Ex.Message);
                                }
                            }
                            break;
                    }
                }
            }
            catch (Exception Ex)
            {
                SBO_Application.StatusBar.SetText(Ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        public static void SetMenuSynchEnable(SAPbouiCOM.Form oForm)
        {
            //Disable all the menus first
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_BP_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_BP_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_ITEM_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_ITEM_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_ITEM_CATEGORY_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_ITEM_CATEGORY_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_ITEM_BARCODES_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_ITEM_BARCODES_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_WAREHOUSE_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_WAREHOUSE_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_BIN_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_BIN_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_BOM_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_BOM_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_BRAND_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_BRAND_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_SALES_ORDER_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_SALES_ORDER_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_RES_INVOICE_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_RES_INVOICE_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_AR_CN_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_AR_CN_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_AR_RETURN_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_AR_RETURN_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_PO_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_PO_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_AP_RETURN_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_AP_RETURN_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_AP_CN_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_AP_CN_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_WORK_ORDER_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_WORK_ORDER_MENU_UID).Enabled = false;
            if (SBO_Application.Menus.Exists(SBOAddon.SYNCH_STOCK_COUNT_MENU_UID)) SBO_Application.Menus.Item(SBOAddon.SYNCH_STOCK_COUNT_MENU_UID).Enabled = false;
            //BoPermission isUserAuthorizeToSynch = eCommon.GetUserAuthorization(SBOAddon.oCompany.UserSignature, String.Format("{0}__AXCSynch", oForm.TypeEx));
            //if (isUserAuthorizeToSynch != BoPermission.boper_Full) return;

            if (oForm.Mode == BoFormMode.fm_OK_MODE)
            {
                switch (oForm.TypeEx)
                {
                    case "134":     //BP Form
                        if (oForm.DataSources.DBDataSources.Item("OCRD").GetValue("U_AXC_SYNCH", 0).ToString() != "Y")
                            SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_BP_MENU_UID).Enabled = true;
                        break;
                    case "139":     //Sales Order
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_SALES_ORDER_MENU_UID).Enabled = true;
                        break;
                    case "142":     //Purchase Order
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_PO_MENU_UID).Enabled = true;
                        break;
                    case "150":     //Items
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_ITEM_MENU_UID).Enabled = true;
                        break;
                    case "1470000002":      //Bin Locations
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_BIN_MENU_UID).Enabled = true;
                        break;
                    case "1470000020":     //BarCodes
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_ITEM_BARCODES_MENU_UID).Enabled = true;
                        break;
                    case "179":     //AR CN
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_AR_CN_MENU_UID).Enabled = true;
                        break;
                    case "180":     //AR Returns
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_AR_RETURN_MENU_UID).Enabled = true;
                        break;
                    case "181":     //AP CN
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_AP_CN_MENU_UID).Enabled = true;
                        break;
                    case "182":     //Purchase Return
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_AP_RETURN_MENU_UID).Enabled = true;
                        break;
                    case "60091":     //AR Reserve Invoice
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_RES_INVOICE_MENU_UID).Enabled = true;
                        break;
                    case "62":      //Warehouses
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_WAREHOUSE_MENU_UID).Enabled = true;
                        break;
                    case "63":     //Item Groups
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_ITEM_CATEGORY_MENU_UID).Enabled = true;
                        break;
                    case "65211":     //Work Order
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_WORK_ORDER_MENU_UID).Enabled = true;
                        break;
                    case "672":        //BOM
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_BOM_MENU_UID).Enabled = true;
                        break;
                    case "897":     //Brands
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_BRAND_MENU_UID).Enabled = true;
                        break;
                    case "1474000001":     //Stock Count
                        SBOAddon.SBO_Application.Menus.Item(SBOAddon.SYNCH_STOCK_COUNT_MENU_UID).Enabled = true;
                        break;
                }
            }
        }

     }
}