using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM;
using AXC_EOA_WMSIntegration;
using AXC_EOA_WMSIntegration.Src.APIAccess;

namespace AXC_EOA_WMSIntegration
{
    [Form("axcOSYNC", true, "Synch To WMS", "", 3)]
    [Authorization("axcOSYNC", "Synch To WMS", "", SAPbobsCOM.BoUPTOptions.bou_FullNone)]
    public class axcOSYNC : Src.Support.UserForm
    {
        SAPbouiCOM.ComboBox _cboOBJTP, _cboFRCODE, _cboTOCODE = null;
        SAPbouiCOM.Button _btnNext = null;
        SAPbouiCOM.Button _btnBack = null;
        SAPbouiCOM.EditText _txtFRCODE = null;
        SAPbouiCOM.EditText _txtTOCODE = null;
        SAPbouiCOM.Grid _grdExport = null;
        SAPbouiCOM.DataTable _dtExport = null;

        public axcOSYNC()
            : base()
        {
        }

        public axcOSYNC(SAPbouiCOM.Form oForm)
            : base(oForm)
        {
        }

        public axcOSYNC(String FormUID)
            : base(FormUID)
        {
        }

        protected override void InitForm()
        {
            _oForm.Freeze(true);
            try
            {
                //Form only for ok mode
                _oForm.SupportedModes = 9;

                //Init Choose From Lists
                InitChooseFromLists();
                

                _cboOBJTP.FillValidValues(Src.Resource.Queries.axcOSYNC_GET_LIST_OBJECTS);
                _cboOBJTP.Select(0, BoSearchKey.psk_Index);     //0 is Group = CompanyDBs


                _oForm.ActiveItem = "cboOBJTP";
                _oForm.PaneLevel = 1;
            }
            finally
            {
                _oForm.Freeze(false);
            }
        }

        protected override void GetItemReferences()
        {
            _cboOBJTP = _oForm.Items.Item("cboOBJTP").Specific as SAPbouiCOM.ComboBox;
            _cboFRCODE = _oForm.Items.Item("cboFRCODE").Specific as SAPbouiCOM.ComboBox;
            _cboTOCODE = _oForm.Items.Item("cboTOCODE").Specific as SAPbouiCOM.ComboBox;
            _grdExport = _oForm.Items.Item("grdExport").Specific as SAPbouiCOM.Grid;
            _dtExport = _oForm.DataSources.DataTables.Item("dtExport");
            _txtFRCODE = _oForm.Items.Item("txtFRCODE").Specific as SAPbouiCOM.EditText;
            _txtTOCODE = _oForm.Items.Item("txtTOCODE").Specific as SAPbouiCOM.EditText;

            _btnNext = _oForm.Items.Item("btnNext").Specific as SAPbouiCOM.Button;
            _btnBack = _oForm.Items.Item("btnBack").Specific as SAPbouiCOM.Button;


            //The Events
            _btnNext.PressedBefore += _btnNext_PressedBefore;
            _btnNext.PressedAfter += _btnNext_PressedAfter;
            _btnBack.PressedAfter += _btnBack_PressedAfter;
            _txtFRCODE.ChooseFromListAfter += _txt_ChooseFromListAfter;
            _txtTOCODE.ChooseFromListAfter += _txt_ChooseFromListAfter;
            _grdExport.LinkPressedBefore += _grdExport_LinkPressedBefore;
            _grdExport.PressedAfter += _grdExport_PressedAfter;
            _cboOBJTP.ComboSelectAfter += _cboOBJTP_ComboSelectAfter;


        }

        void _txt_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.EditText edt = sboObject as SAPbouiCOM.EditText;
                SAPbouiCOM.ISBOChooseFromListEventArg pCFL = pVal as SAPbouiCOM.ISBOChooseFromListEventArg;

                if (pCFL.SelectedObjects != null)
                {
                    _oForm.DataSources.UserDataSources.Item(edt.DataBind.Alias).ValueEx = pCFL.SelectedObjects.GetValue(edt.ChooseFromListAlias, 0).ToString();
                }
            }
            catch (Exception Ex)
            {
                SBOAddon.SBO_Application.StatusBar.SetText(Ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        void _cboOBJTP_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                _oForm.Freeze(true);
                string selectedObj = _oForm.DataSources.UserDataSources.Item("cboOBJTP").ValueEx;
                _txtFRCODE.Value = "";
                _txtTOCODE.Value = "";
                _oForm.DataSources.UserDataSources.Item(_cboFRCODE.DataBind.Alias).ValueEx = "";
                _oForm.DataSources.UserDataSources.Item(_cboTOCODE.DataBind.Alias).ValueEx = "";
                _oForm.ActiveItem = "cboOBJTP";

                _txtFRCODE.Item.Visible = selectedObj != Brands._objectType;
                _txtTOCODE.Item.Visible = selectedObj != Brands._objectType;
                _cboFRCODE.Item.Visible = selectedObj == Brands._objectType;
                _cboTOCODE.Item.Visible = selectedObj == Brands._objectType;
                switch (selectedObj)
                {
                    case ARCNotes._objectType: SetARCNChooseFromList(true); break;
                    case APCNotes._objectType: SetAPCNChooseFromList(true); break;
                    case APReturns._objectType: SetAPReturnChooseFromList(true); break;
                    case ARResInvoices._objectType: SetRInvoiceChooseFromList(true); break;
                    case ARReturns._objectType: SetARReturnChooseFromList(true); break;
                    case Barcodes._objectType: SetBarCodeChooseFromList(true); break;
                    case BillOfMaterials._objectType: SetBOMChooseFromList(true); break;
                    case BinLocations._objectType: SetBinChooseFromList(true); break;
                    case Brands._objectType: SetBrandComboBox(true); break;
                    case Customer._objectType: SetCustomerChooseFromList(true); break;
                    case ItemCategories._objectType: SetItemCatChooseFromList(true); break;
                    case Src.APIAccess.Items._objectType: SetItemChooseFromList(true); break;
                    case PurchaseOrders._objectType: SetPOChooseFromList(true); break;
                    case SalesOrders._objectType: SetSOChooseFromList(true); break;
                    case Vendor._objectType: SetVendorChooseFromList(true); break;
                    case Warehouses._objectType: SetWarehouseChooseFromList(true); break;
                    case WorkOrders._objectType: SetWorkOrderChooseFromList(true); break;
                    case ItemCount._objectType: SetItemCountChooseFromList(true); break;
                }
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

        void _grdExport_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {

            try
            {
                _oForm.Freeze(true);
                switch (pVal.ColUID)
                {
                    case "Export":
                        if (pVal.Row == -1)
                        {
                            string select = "Y";
                            if (_dtExport.GetValue("Export", 0).ToString().Trim() == "Y") select = "N";

                            for (int rowIdx = 0; rowIdx < _dtExport.Rows.Count; rowIdx++)
                            {
                                _dtExport.SetValue("Export", rowIdx, select);
                            }
                        
                        }
                        break;
                }
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

        void _grdExport_LinkPressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.ColUID)
                {
                    case "Code":
                    case "Group Name":
                    case "DocNum":
                    case "BinCode":
                        String ObjType = _dtExport.GetValue("ObjType", _grdExport.GetDataTableRowIndex(pVal.Row)).ToString().Trim();
                        String Code = _dtExport.GetValue("Key", _grdExport.GetDataTableRowIndex(pVal.Row)).ToString().Trim();
                        int iObjType = 0;
                        if (int.TryParse(ObjType, out iObjType))
                        {
                            SBOAddon.SBO_Application.OpenForm((BoFormObjectEnum)iObjType, "", Code);
                        }
                        BubbleEvent = false;
                        break;
                    case "SalesOrder":  //Sales Order as the origin of a work orders
                        if(_cboOBJTP.Selected.Value == WorkOrders._objectType)
                        {
                            string soEntry = _dtExport.GetValue("OriginAbs", _grdExport.GetDataTableRowIndex(pVal.Row)).ToString().Trim();
                            SBOAddon.SBO_Application.OpenForm(BoFormObjectEnum.fo_Order, "", soEntry);
                            BubbleEvent = false;
                        }
                        break;
                }
            }
            catch (Exception Ex)
            {
                SBOAddon.SBO_Application.StatusBar.SetText(Ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }

        }

        void _btnBack_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                _oForm.PaneLevel = 1;
            }
            catch (Exception Ex)
            {
                SBOAddon.SBO_Application.StatusBar.SetText(Ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        void _btnNext_PressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (_oForm.PaneLevel)
                {
                    case 1: 
                        ValidatePane1();
                        break;
                    case 2:
                        ValidatePane2();
                        break;
                }

            }
            catch(Exception ex)
            {
                SBOAddon.SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        void _btnNext_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                switch(_oForm.PaneLevel)
                {
                    case 1:
                        PreparePane2();
                        break;
                    case 2:
                        SendToWMS();
                        break;
                }
            }
            catch (Exception ex)
            {
                SBOAddon.SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }


        [FormEvent("ResizeAfter", false)]
        public static void OnAfterFormResize(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = AXC_EOA_WMSIntegration.Src.Support.Addon.SBO_Application.Forms.Item(pVal.FormUID);
                if (oForm.Items.Count > 0)
                {
                    //oForm.Items.Item("Item_16").Width = oForm.ClientWidth - 10;
                    //oForm.Items.Item("Item_16").Height = oForm.Items.Item("btnNext").Top - 10 - oForm.Items.Item("Item_16").Top;
                }
            }
            catch (Exception Ex)
            {
                SBOAddon.SBO_Application.StatusBar.SetText(Ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        protected override void OnBeforeFormClose(SBOItemEventArg pVal, out bool Bubble)
        {
            Bubble = true;
            try
            {
                //Remove all the events
                _btnNext.PressedBefore -= _btnNext_PressedBefore;
                _btnBack.PressedAfter -= _btnBack_PressedAfter;

            }
            catch { };
        }

        private void FormatGridExport()
        {
            for (int iCol = 0; iCol < _grdExport.Columns.Count; iCol++)
            {
                _grdExport.Columns.Item(iCol).Editable = false;
                _grdExport.Columns.Item(iCol).TitleObject.Sortable = true;
                _grdExport.Columns.Item(iCol).RightJustified = _dtExport.Columns.Item(iCol).Type != BoFieldsType.ft_AlphaNumeric 
                    && _dtExport.Columns.Item(iCol).Type != BoFieldsType.ft_Date
                    && _dtExport.Columns.Item(iCol).Type != BoFieldsType.ft_NotDefined
                    && _dtExport.Columns.Item(iCol).Type != BoFieldsType.ft_Text;
                SAPbouiCOM.EditTextColumn edtCol = null;
                switch (_grdExport.Columns.Item(iCol).UniqueID)
                {
                    case "Export":
                        _grdExport.Columns.Item(iCol).Type = BoGridColumnType.gct_CheckBox;
                        _grdExport.Columns.Item(iCol).Editable = true;
                        break;
                    case "Code":
                    case "DocNum":
                    case "Group Name":
                        edtCol = _grdExport.Columns.Item(iCol) as SAPbouiCOM.EditTextColumn;
                        if (_cboOBJTP.Selected.Value == "")
                            edtCol.LinkedObjectType = "";
                        else if(_cboOBJTP.Selected.Value == "13R")
                            edtCol.LinkedObjectType = ((int)SAPbobsCOM.BoObjectTypes.oInvoices).ToString();
                        else
                            edtCol.LinkedObjectType = _cboOBJTP.Selected.Value;    //Dummy to show the link button
                        break;
                    case "WhsCode":
                        edtCol = _grdExport.Columns.Item(iCol) as SAPbouiCOM.EditTextColumn;
                        edtCol.LinkedObjectType = ((int)SAPbobsCOM.BoObjectTypes.oWarehouses).ToString();
                        break;
                    case "ItemCode":
                        edtCol = _grdExport.Columns.Item(iCol) as SAPbouiCOM.EditTextColumn;
                        edtCol.LinkedObjectType = ((int)SAPbobsCOM.BoObjectTypes.oItems).ToString();
                        break;
                    case "Preferred Vendor":
                    case "CardCode":
                        edtCol = _grdExport.Columns.Item(iCol) as SAPbouiCOM.EditTextColumn;
                        edtCol.LinkedObjectType = ((int)SAPbobsCOM.BoObjectTypes.oBusinessPartners).ToString();
                        break;
                    case "BinCode":
                        SAPbouiCOM.EditTextColumn oCol = _grdExport.Columns.Item(iCol) as SAPbouiCOM.EditTextColumn;
                        if (_cboOBJTP.Selected.Value == "")
                            oCol.LinkedObjectType = "";
                        else
                            oCol.LinkedObjectType = "1";    //Dummy to show the link button
                        break;
                    case "SalesOrder":
                        edtCol = _grdExport.Columns.Item(iCol) as SAPbouiCOM.EditTextColumn;
                        edtCol.LinkedObjectType = ((int)SAPbobsCOM.BoObjectTypes.oOrders).ToString();
                        break;
                    case "Active":
                    case "Inactive":
                    case "Bin Warehouse":
                    case "Success":
                    case "Locked":
                    case "ReceiveBin":
                    case "Disabled":
                        _grdExport.Columns.Item(iCol).Type = BoGridColumnType.gct_CheckBox;
                        break;
                    case "OriginAbs":   
                    case "Key":
                    case "ObjType":
                    case "GroupCode":
                    case "Name":
                        _grdExport.Columns.Item(iCol).Visible = false;
                        break;

                }
            }
            _grdExport.AutoResizeColumns();
        }

        private void InitChooseFromLists()
        {
            //Choose From Lists

            //public const string SYNCH_O_OBJECT_VENDOR = "2V";
            SetVendorChooseFromList(attach: false);
            //public const string SYNCH_O_OBJECT_ITEM = "4";
            SetItemChooseFromList(attach: false);
            //public const string SYNCH_O_OBJECT_ITEM_CATEGORY = "8";
            SetItemCatChooseFromList(attach: false);
            //public const string SYNCH_O_OBJECT_BRAND = "43";
            SetBrandComboBox(attach: false);
            //public const string SYNCH_O_OBJECT_BAR_CODE = "1470000062";
            SetBarCodeChooseFromList(attach: false);
            //public const string SYNCH_O_OBJECT_WAREHOUSE = "64";
            SetWarehouseChooseFromList(attach: false);
            //public const string SYNCH_O_OBJECT_BIN = "10000206";
            SetBinChooseFromList(attach: false);
            //public const string SYNCH_O_OBJECT_BOM = "66";
            SetBOMChooseFromList(attach: false);
            //public const string SYNCH_O_OBJECT_SALES_ORDER = "17";
            SetSOChooseFromList(attach: false);
            //public const string SYNCH_O_OBJECT_RESERVE_INVOICE = "13R";
            SetRInvoiceChooseFromList(attach: false);
            //public const string SYNCH_O_OBJECT_AR_CN = "14";
            SetARCNChooseFromList(attach: false);
            //public const string SYNCH_O_OBJECT_AR_RETURNS = "16";
            SetARReturnChooseFromList(attach: false);
            //public const string SYNCH_O_OBJECT_PURCHASE_ORDER = "22";
            SetPOChooseFromList(attach: false);
            //public const string SYNCH_O_OBJECT_AP_RETURN = "21";
            SetAPReturnChooseFromList(attach: false);
            //public const string SYNCH_O_OBJECT_AP_CN = "19";
            SetAPCNChooseFromList(attach: false);
            //public const string SYNCH_O_OBJECT_WORK_ORDER = "202";
            SetWorkOrderChooseFromList(attach: false);
            //public const string SYNCH_O_OBJECT_STOCK_COUNT = "1470000065";
            SetItemCountChooseFromList(attach: false);


        }

        private void SetWorkOrderChooseFromList(bool attach)
        {
            if(attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFOWOR";
                _txtTOCODE.ChooseFromListUID = "cflTOWOR";
                _txtFRCODE.ChooseFromListAlias = WorkOrders._filterField;
                _txtTOCODE.ChooseFromListAlias = WorkOrders._filterField;
            }            
        }

        private void SetItemCountChooseFromList(bool attach)
        {
            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFOINC";
                _txtTOCODE.ChooseFromListUID = "cflTOINC";
                _txtFRCODE.ChooseFromListAlias = ItemCount._filterField;
                _txtTOCODE.ChooseFromListAlias = ItemCount._filterField;
            }
        }

        private void SetAPCNChooseFromList(bool attach)
        {
            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFORPC";
                _txtTOCODE.ChooseFromListUID = "cflTORPC";
                _txtFRCODE.ChooseFromListAlias = APCNotes._filterField;
                _txtTOCODE.ChooseFromListAlias = APCNotes._filterField;
            }
        }

        private void SetAPReturnChooseFromList(bool attach)
        {
            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFORPD";
                _txtTOCODE.ChooseFromListUID = "cflTORPD";
                _txtFRCODE.ChooseFromListAlias = APReturns._filterField;
                _txtTOCODE.ChooseFromListAlias = APReturns._filterField;
            }
        }

        private void SetPOChooseFromList(bool attach)
        {
            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFOPOR";
                _txtTOCODE.ChooseFromListUID = "cflTOPOR";
                _txtFRCODE.ChooseFromListAlias = PurchaseOrders._filterField;
                _txtTOCODE.ChooseFromListAlias = PurchaseOrders._filterField;
            }
        }

        private void SetARReturnChooseFromList(bool attach)
        {
            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFORDN";
                _txtTOCODE.ChooseFromListUID = "cflTORDN";
                _txtFRCODE.ChooseFromListAlias = ARReturns._filterField;
                _txtTOCODE.ChooseFromListAlias = ARReturns._filterField;
            }
        }

        private void SetARCNChooseFromList(bool attach)
        {
            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFORIN";
                _txtTOCODE.ChooseFromListUID = "cflTORIN";
                _txtFRCODE.ChooseFromListAlias = ARCNotes._filterField;
                _txtTOCODE.ChooseFromListAlias = ARCNotes._filterField;
            }
        }

        private void SetRInvoiceChooseFromList(bool attach)
        {
            //Reserve Invoices
            SAPbouiCOM.Conditions conds = new SAPbouiCOM.Conditions();
            var cond = conds.Add();
            //T0.[IsICT] = 'N'  	AND  T0.[UpdInvnt] = 'C'  	AND  T0.[DocSubType] = N'--'
            cond.Alias = "IsICT";
            cond.Operation = BoConditionOperation.co_EQUAL;
            cond.CondVal = "N";
            cond.Relationship = BoConditionRelationship.cr_AND;

            cond = conds.Add();
            cond.Alias = "UpdInvnt";
            cond.Operation = BoConditionOperation.co_EQUAL;
            cond.CondVal = "C";
            cond.Relationship = BoConditionRelationship.cr_AND;

            cond = conds.Add();
            cond.Alias = "DocSubType";
            cond.Operation = BoConditionOperation.co_EQUAL;
            cond.CondVal = "--";

            _oForm.ChooseFromLists.Item("cflFOINV").SetConditions(conds);
            _oForm.ChooseFromLists.Item("cflTOINV").SetConditions(conds);

            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFOINV";
                _txtTOCODE.ChooseFromListUID = "cflTOINV";
                _txtFRCODE.ChooseFromListAlias = ARResInvoices._filterField;
                _txtTOCODE.ChooseFromListAlias = ARResInvoices._filterField;
            }
        }

        private void SetSOChooseFromList(bool attach)
        {
            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFORDR";
                _txtTOCODE.ChooseFromListUID = "cflTORDR";
                _txtFRCODE.ChooseFromListAlias = SalesOrders._filterField;
                _txtTOCODE.ChooseFromListAlias = SalesOrders._filterField;
            }
        }

        private void SetBOMChooseFromList(bool attach)
        {
            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFOITT";
                _txtTOCODE.ChooseFromListUID = "cflTOITT";
                _txtFRCODE.ChooseFromListAlias = BillOfMaterials._filterField;
                _txtTOCODE.ChooseFromListAlias = BillOfMaterials._filterField;
            }
        }

        private void SetBinChooseFromList(bool attach)
        {
            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFOBIN";
                _txtTOCODE.ChooseFromListUID = "cflTOBIN";
                _txtFRCODE.ChooseFromListAlias = BinLocations._filterField;
                _txtTOCODE.ChooseFromListAlias = BinLocations._filterField;
            }
        }

        private void SetWarehouseChooseFromList(bool attach)
        {
            //DropShip = 'N'
            SAPbouiCOM.Conditions conds = new SAPbouiCOM.Conditions();
            var cond = conds.Add();
            cond.Alias = "DropShip";
            cond.Operation = BoConditionOperation.co_EQUAL;
            cond.CondVal = "N";
            _oForm.ChooseFromLists.Item("cflFOWHS").SetConditions(conds);
            _oForm.ChooseFromLists.Item("cflTOWHS").SetConditions(conds);

            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFOWHS";
                _txtTOCODE.ChooseFromListUID = "cflTOWHS";
                _txtFRCODE.ChooseFromListAlias = Warehouses._filterField;
                _txtTOCODE.ChooseFromListAlias = Warehouses._filterField;
            }
        }

        private void SetBarCodeChooseFromList(bool attach)
        {
            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFOBCD";
                _txtTOCODE.ChooseFromListUID = "cflTOBCD";
                _txtFRCODE.ChooseFromListAlias = Barcodes._filterField;
                _txtTOCODE.ChooseFromListAlias = Barcodes._filterField;
            }
        }

        private void SetBrandComboBox(bool attach)
        {
            if (attach)
            {
                _cboFRCODE.FillValidValues(Src.Resource.Queries.OMRC_GET_FIRM_NAMES);
                _cboTOCODE.FillValidValues(Src.Resource.Queries.OMRC_GET_FIRM_NAMES);
                _cboFRCODE.Select("0", BoSearchKey.psk_ByValue);
                _cboTOCODE.Select("0", BoSearchKey.psk_ByValue);
            }
        }

        private void SetItemCatChooseFromList(bool attach)
        {
            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFOITB";
                _txtTOCODE.ChooseFromListUID = "cflTOITB";
                _txtFRCODE.ChooseFromListAlias = ItemCategories._filterField;
                _txtTOCODE.ChooseFromListAlias = ItemCategories._filterField;
            }
        }

        private void SetItemChooseFromList(bool attach)
        {
            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFOITM";
                _txtTOCODE.ChooseFromListUID = "cflTOITM";
                _txtFRCODE.ChooseFromListAlias = Src.APIAccess.Items._filterField;
                _txtTOCODE.ChooseFromListAlias = Src.APIAccess.Items._filterField;
            }
        }

        private void SetCustomerChooseFromList(bool attach)
        {
            string cardType = "C";
            SAPbouiCOM.Conditions conds = new SAPbouiCOM.Conditions();
            var cond = conds.Add();
            cond.Alias = "CardType";
            cond.Operation = BoConditionOperation.co_EQUAL;
            cond.CondVal = cardType;
            _oForm.ChooseFromLists.Item("cflFOCRD").SetConditions(conds);
            _oForm.ChooseFromLists.Item("cflTOCRD").SetConditions(conds);

            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFOCRD";
                _txtTOCODE.ChooseFromListUID = "cflTOCRD";
                _txtFRCODE.ChooseFromListAlias = Customer._filterField;
                _txtTOCODE.ChooseFromListAlias = Customer._filterField;
            }
        }

        private void SetVendorChooseFromList(bool attach)
        {
            string cardType = "S";
            SAPbouiCOM.Conditions conds = new SAPbouiCOM.Conditions();
            var cond = conds.Add();
            cond.Alias = "CardType";
            cond.Operation = BoConditionOperation.co_EQUAL;
            cond.CondVal = cardType;
            _oForm.ChooseFromLists.Item("cflFOCRD").SetConditions(conds);
            _oForm.ChooseFromLists.Item("cflTOCRD").SetConditions(conds);

            if (attach)
            {
                _txtFRCODE.ChooseFromListUID = "cflFOCRD";
                _txtTOCODE.ChooseFromListUID = "cflTOCRD";
                _txtFRCODE.ChooseFromListAlias = Vendor._filterField;
                _txtTOCODE.ChooseFromListAlias = Vendor._filterField;
            }
        }

        private void PreparePane2()
        {
            _oForm.Freeze(true);
            try
            {
                //Query the grids and set the pane level
                String objectType = _cboOBJTP.Selected?.Value.Trim().ToUpper()??"%";
                String sSQLExport = "";
                String fromCode = _oForm.DataSources.UserDataSources.Item("txtFRCODE").ValueEx.Trim();
                String toCode = _oForm.DataSources.UserDataSources.Item("txtTOCODE").ValueEx.Trim();
                if(objectType ==  Src.APIAccess.Brands._objectType)
                {
                    fromCode = _cboFRCODE.Selected?.Description??"";
                    toCode = _cboTOCODE.Selected?.Description??"";
                }

                if (_cboOBJTP.Selected != null)
                {
                    if (fromCode == "") fromCode = toCode;
                    if (toCode == "") toCode = fromCode;

                    switch (objectType)
                    {
                        case Src.APIAccess.APCNotes._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_ORPC, fromCode == "" ? "0" : fromCode, toCode == "" ? "0" : toCode);
                            break;
                        case Src.APIAccess.APReturns._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_ORPD, fromCode == "" ? "0" : fromCode, toCode == "" ? "0" : toCode);
                            break;
                        case Src.APIAccess.ARCNotes._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_ORIN, fromCode == "" ? "0" : fromCode, toCode == "" ? "0" : toCode);
                            break;
                        case Src.APIAccess.ARResInvoices._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_OINV_RESERVE, fromCode == "" ? "0" : fromCode, toCode == "" ? "0" : toCode);
                            break;
                        case Src.APIAccess.ARReturns._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_ORDN, fromCode == "" ? "0" : fromCode, toCode == "" ? "0" : toCode);
                            break;
                        case Src.APIAccess.PurchaseOrders._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_OPOR, fromCode == "" ? "0" : fromCode, toCode == "" ? "0" : toCode);
                            break;
                        case Src.APIAccess.SalesOrders._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_ORDR, fromCode == "" ? "0" : fromCode, toCode == "" ? "0" : toCode);
                            break;
                        case Src.APIAccess.WorkOrders._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_OWOR, fromCode == "" ? "0" : fromCode, toCode == "" ? "0" : toCode);
                            break;
                        case Src.APIAccess.Barcodes._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_OBCD, fromCode, toCode);
                            break;
                        case Src.APIAccess.BillOfMaterials._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_OITT, fromCode, toCode);
                            break;
                        case Src.APIAccess.BinLocations._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_OBIN, fromCode, toCode);
                            break;
                        case Src.APIAccess.Brands._objectType:           //Items
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_OMRC, fromCode, toCode);
                            break;
                        case Src.APIAccess.Vendor._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_VENDORS, fromCode, toCode);
                            break;
                        case Src.APIAccess.Customer._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_CUSTOMERS, fromCode, toCode);
                            break;
                        case Src.APIAccess.ItemCategories._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_OITB, fromCode, toCode);
                            break;
                        case Src.APIAccess.Items._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_OITM, fromCode, toCode);
                            break;
                        case Src.APIAccess.Warehouses._objectType:
                            sSQLExport = String.Format(Src.Resource.Queries.axcOSYNC_GET_LIST_OWHS, fromCode, toCode);
                            break;
                    }
                }

                _dtExport.ExecuteQuery(sSQLExport);
                _grdExport.DataTable = _dtExport;
                FormatGridExport();

                _oForm.PaneLevel = 2;
            }
            catch (Exception Ex)
            {
                _oForm.PaneLevel = 1;
                throw;
            }
            finally
            {
                _oForm.Freeze(false);
            }
        }

        private void SendToWMS()
        {
            SAPbouiCOM.ProgressBar pb = null;
            try
            {
                _oForm.Freeze(true);

                int[] selectedRows = _dtExport.DataTableIndexOf("Export", "Y");
                if (selectedRows == null)
                    return;
                pb = eCommon.TryCreateProgressBar("Sending objects", selectedRows.Length, true);
                pb.TrySetValue("", 0);

                string objCode = _cboOBJTP.Selected.Value;
                bool allSuccess = true;
                foreach (int row in selectedRows)
                {
                    pb.TrySetValue();
                    string result = "";
                    bool success = false;
                    try
                    {
                        string objectName = _dtExport.GetValue("Name", row).ToString();
                        object objectKey;

                        switch (objCode)
                        {
                            case Src.APIAccess.APCNotes._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Src.APIAccess.APCNotes>(objectKey, objectName, out result);
                                break;
                            case Src.APIAccess.APReturns._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Src.APIAccess.APReturns>(objectKey, objectName, out result);
                                break;
                            case Src.APIAccess.ARCNotes._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Src.APIAccess.ARCNotes>(objectKey, objectName, out result);
                                break;
                            case Src.APIAccess.ARResInvoices._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Src.APIAccess.ARResInvoices>(objectKey, objectName, out result);
                                break;
                            case Src.APIAccess.ARReturns._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Src.APIAccess.ARReturns>(objectKey, objectName, out result);
                                break;
                            case Src.APIAccess.Barcodes._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Src.APIAccess.Barcodes>(objectKey, objectName, out result);
                                break;
                            case Src.APIAccess.BillOfMaterials._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Src.APIAccess.BillOfMaterials>(objectKey, objectName, out result);
                                break;
                            case Src.APIAccess.BinLocations._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Src.APIAccess.BinLocations>(objectKey, objectName, out result);
                                break;
                            case Src.APIAccess.Brands._objectType:           //Items
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Src.APIAccess.Brands>(objectKey, objectName, out result);
                                break;
                            case Src.APIAccess.Vendor._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Vendor>(objectKey.ToString(), objectName, out result);
                                break;
                            case Src.APIAccess.Customer._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Customer>(objectKey.ToString(), objectName, out result);
                                break;
                            case Src.APIAccess.ItemCategories._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Src.APIAccess.ItemCategories>(objectKey, objectName, out result);
                                break;
                            case Src.APIAccess.Items._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Src.APIAccess.Items>(objectKey, objectName, out result);
                                break;
                            case Src.APIAccess.PurchaseOrders._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Src.APIAccess.PurchaseOrders>(objectKey, objectName, out result);
                                break;
                            case Src.APIAccess.SalesOrders._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Src.APIAccess.SalesOrders>(objectKey, objectName, out result);
                                break;
                            case Src.APIAccess.Warehouses._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Src.APIAccess.Warehouses>(objectKey, objectName, out result);
                                break;
                            case Src.APIAccess.WorkOrders._objectType:
                                objectKey = _dtExport.GetValue("Key", row);
                                success = APIServiceAccess.SynchObject<Src.APIAccess.WorkOrders>(objectKey, objectName, out result);
                                break;
                        }
                    }
                    catch (Exception Ex)
                    {
                        success = false;
                        result = Ex.Message;
                    }
                    finally
                    {
                        if (!success) allSuccess = false;
                        //Update each row 
                        _dtExport.SetValue("Success", row, success ? "Y" : "N");

                        if (result.Length > 254)
                            result = result.Substring(0, 254);
                        _dtExport.SetValue("Export Process Remark", row, result);

                    }
                }
                if (allSuccess)
                    SBOAddon.SBO_Application.StatusBar.SetText("Operation completed successfully.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                else
                    SBOAddon.SBO_Application.StatusBar.SetText("Operation completed with error. Check each line for problem if exists.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

            }
            finally
            {
                _oForm.Freeze(false);
                pb.TryStop();
            }
        }

        private void ValidatePane1()
        {
            return;
        }

        private void ValidatePane2()
        {
            int[] selectedRows = _dtExport.DataTableIndexOf("Export", "Y");
            if (selectedRows == null || selectedRows.Length == 0)
                throw new Exception("Please select at least 1 row to continue.");

            int userReply = SBOAddon.SBO_Application.MessageBox($"Send {selectedRows.Length} records to WMS?", 2, "Yes", "Cancel");
            if (userReply != 1)
                throw new Exception("User cancelled operation.");

        }
    }
}


