using System;
using System.Collections.Generic;
using System.Text;
using AXC_EOA_WMSIntegration.Src.APIAccess;
using SAPbouiCOM;

namespace AXC_EOA_WMSIntegration
{
    [Form("axcOFTLG", true, "Integration Log", "", 2)]
    [Authorization("axcOFTLG", "Integration Log", "", SAPbobsCOM.BoUPTOptions.bou_FullNone)]
    public class axcOFTLG : Src.Support.UserForm
    {
        SAPbouiCOM.EditText _txtFRDTE = null;
        SAPbouiCOM.EditText _txtTODTE = null;
        SAPbouiCOM.ComboBox _cboOBJTP = null;
        SAPbouiCOM.ComboBox _cboUSER = null;
        SAPbouiCOM.ComboBox _cboSCCES = null;
        SAPbouiCOM.ComboBox _cboDRCTN = null;
        SAPbouiCOM.Button _btnNext = null;
        SAPbouiCOM.Button _btnBack = null;

        SAPbouiCOM.Grid _grdLog = null;
        SAPbouiCOM.DataTable _dtLog = null;


        public axcOFTLG()
            : base()
        {
        }

        public axcOFTLG(SAPbouiCOM.Form oForm)
            : base(oForm)
        {
        }

        public axcOFTLG(String FormUID)
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
                
                _cboUSER.FillValidValues(Src.Resource.Queries.OFTLOG_GET_COMBO_BOX_USERS);
                _cboOBJTP.FillValidValues( Src.Resource.Queries.OFTLOG_GET_COMBO_BOX_OBJECTS);
                _cboDRCTN.FillValidValues(Src.Resource.Queries.OFTLOG_GET_COMBO_BOX_DIRECTION);

                _cboUSER.Select(0, BoSearchKey.psk_Index);
                _cboOBJTP.Select(0, BoSearchKey.psk_Index);
                _cboSCCES.Select(0, BoSearchKey.psk_Index);
                _cboDRCTN.Select(0, BoSearchKey.psk_Index);

                _oForm.ActiveItem = _txtFRDTE.Item.UniqueID;
                _oForm.PaneLevel = 1;
            }
            catch (Exception Ex)
            {
                AXC_EOA_WMSIntegration.Src.Support.Addon.SBO_Application.MessageBox(Ex.Message);
            }
            finally
            {
                _oForm.Freeze(false);
            }
        }

        protected override void GetItemReferences()
        {
            try
            {
                _txtFRDTE = _oForm.Items.Item("txtFRDTE").Specific as SAPbouiCOM.EditText;
                _txtTODTE = _oForm.Items.Item("txtTODTE").Specific as SAPbouiCOM.EditText;
                _cboOBJTP = _oForm.Items.Item("cboOBJTP").Specific as SAPbouiCOM.ComboBox;
                _cboUSER = _oForm.Items.Item("cboUSER").Specific as SAPbouiCOM.ComboBox;
                _cboSCCES = _oForm.Items.Item("cboSCCES").Specific as SAPbouiCOM.ComboBox;
                _cboDRCTN = _oForm.Items.Item("cboDRCTN").Specific as SAPbouiCOM.ComboBox;
                _grdLog = _oForm.Items.Item("grdLog").Specific as SAPbouiCOM.Grid;
                _dtLog = _oForm.DataSources.DataTables.Item("dtLog");


                _btnNext = _oForm.Items.Item("btnNext").Specific as SAPbouiCOM.Button;
                _btnBack = _oForm.Items.Item("btnBack").Specific as SAPbouiCOM.Button;


                //The Events
                _btnNext.PressedBefore += _btnNext_PressedBefore;
                _btnBack.PressedAfter += _btnBack_PressedAfter;
                _grdLog.LinkPressedBefore += _grdLog_LinkPressedBefore;

            }
            catch (Exception Ex)
            {
                AXC_EOA_WMSIntegration.Src.Support.Addon.SBO_Application.MessageBox(Ex.Message);
            }
        }

        void _grdLog_LinkPressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.ColUID)
                {
                    case "Code":
                        String ObjType = _dtLog.GetValue("ObjType", _grdLog.GetDataTableRowIndex(pVal.Row)).ToString().Trim();
                        String Code = _dtLog.GetValue("Code", _grdLog.GetDataTableRowIndex(pVal.Row)).ToString().Trim();
                        if(ObjType == "43")
                        {
                            SBOAddon.SBO_Application.StatusBar.SetText("Barcode is not accessible from link button", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                            BubbleEvent = false;
                            return;
                        }
                        int iObjType = 0;
                        if (int.TryParse(ObjType, out iObjType))
                        {
                            SBOAddon.SBO_Application.OpenForm((BoFormObjectEnum)iObjType, "", Code);
                        }
                        BubbleEvent = false;
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
            _oForm.Freeze(true);
            try
            {
                //Query the grids and set the pane level
                String FromDate = DateTime.FromOADate(0).ToString("yyyyMMdd");
                String ToDate = "21991231";
                String objectType = "%";
                String Success = "%";
                String userCode = "%";
                String direction = "%";

                if (_txtFRDTE.String.Trim() != "" && _txtTODTE.String.Trim() != "")
                {
                    FromDate = _oForm.DataSources.UserDataSources.Item("txtFRDTE").ValueEx;
                    ToDate = _oForm.DataSources.UserDataSources.Item("txtTODTE").ValueEx;
                }
                else if (_txtFRDTE.String.Trim() == "" && _txtTODTE.String.Trim() == "")
                {
                    FromDate = DateTime.FromOADate(0).ToString("yyyyMMdd");
                    ToDate = "21991231";
                }
                else if (_txtFRDTE.String.Trim() == "")
                {
                    ToDate = _oForm.DataSources.UserDataSources.Item("txtTODTE").Value;
                    FromDate = ToDate;
                }
                else if (_txtTODTE.String.Trim() == "")
                {
                    FromDate = _oForm.DataSources.UserDataSources.Item("txtFRDTE").Value;
                    ToDate = FromDate;
                }

                objectType = _cboOBJTP.Selected?.Value ?? "%";
                Success = _cboSCCES.Selected?.Value ?? "%";
                userCode = _cboUSER.Selected?.Value ?? "%";
                direction = _cboDRCTN.Selected?.Value??"%";

                String sSQL = String.Format(Src.Resource.Queries.OFTLOG_GET_LOG, objectType, Success, FromDate, ToDate, userCode, direction);
                _dtLog.ExecuteQuery(sSQL);
                _grdLog.DataTable = _dtLog;
                FormatGridLog();

                _oForm.PaneLevel = 2;
            }
            catch (Exception Ex)
            {
                _oForm.PaneLevel = 1;
                SBOAddon.SBO_Application.StatusBar.SetText(Ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                _oForm.Freeze(false);
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
                }
            }
            catch(Exception ex)
            {
                SBOAddon.SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            catch { }

        }

        public static void GenerateRecord(String ObjType, String ObjCode, String ObjName, String ExtenalID, APIServiceAccess.Operation Ops, String Data, String Result, Boolean Success)
        {
            SAPbobsCOM.UserTable ut = SBOAddon.oCompany.UserTables.Item(SBOAddon_DB.OFTLG_TABLE_UID.Substring(1));

            ut.UserFields.Fields.Item(SBOAddon_DB.OFTLG_USER_ID).Value = SBOAddon.oCompany.UserSignature;
            ut.UserFields.Fields.Item(SBOAddon_DB.OFTLG_WS_OBJECT_TYPE).Value = ObjType??"";
            ut.UserFields.Fields.Item(SBOAddon_DB.OFTLG_WS_OBJECT_CODE).Value = ObjCode??"";
            ut.UserFields.Fields.Item(SBOAddon_DB.OFTLG_WS_OBJECT_NAME).Value = ObjName??"";
            ut.UserFields.Fields.Item(SBOAddon_DB.OFTLG_WS_OPERATION).Value = Enum.GetName(typeof(APIServiceAccess.Operation), Ops);
            ut.UserFields.Fields.Item(SBOAddon_DB.OFTLF_WS_DIRECTION).Value = "O";      //Outbound
            //if (Data.Length > 4000) Data = Data.Substring(0, 4000);
            ut.UserFields.Fields.Item(SBOAddon_DB.OFTLG_WS_POST_DATA).Value = Data;
            //if (Result.Length > 4000) Result = Result.Substring(0, 4000);
            ut.UserFields.Fields.Item(SBOAddon_DB.OFTLG_WS_POST_RESULT).Value = Result;
            ut.UserFields.Fields.Item(SBOAddon_DB.OFTLG_WS_POST_SUCCESS).Value = Success ? "Y" : "N";
            string currentTime = eCommon.ExecuteScalar(Src.Resource.Queries.axcOFTLG_GET_SERVER_TIME).ToString();
            ut.UserFields.Fields.Item(SBOAddon_DB.OFTLG_WS_EXPORT_TIME_STAMP).Value = currentTime;
            ut.UserFields.Fields.Item(SBOAddon_DB.OFTLG_WS_EXTERNAL_KEY).Value = ExtenalID ?? "";

            int err = ut.Add();
            if (err != 0)
                SBOAddon.SBO_Application.StatusBar.SetText($"Could not create log. {SBOAddon.oCompany.GetLastErrorDescription()}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

            eCommon.ReleaseComObject(ut);
        }

        private void FormatGridLog()
        {
            for (int iCol = 0; iCol < _grdLog.Columns.Count; iCol++)
            {
                _grdLog.Columns.Item(iCol).Editable = false;
                _grdLog.Columns.Item(iCol).TitleObject.Sortable = true;
                SAPbouiCOM.ComboBoxColumn cboCol = null;
                switch (_grdLog.Columns.Item(iCol).UniqueID)
                {
                    case "ID":
                        break;
                    case "TimeStamp":
                        break;
                    case "User":
                        break;
                    case "WMSObjectType":
                        _grdLog.Columns.Item(iCol).Type = BoGridColumnType.gct_ComboBox;
                        cboCol = (SAPbouiCOM.ComboBoxColumn)_grdLog.Columns.Item(iCol);
                        cboCol.FillValidValues(Src.Resource.Queries.OFTLOG_GET_COMBO_BOX_OBJECTS);
                        cboCol.DisplayType = BoComboDisplayType.cdt_both;
                        break;
                    case "Code":
                        SAPbouiCOM.EditTextColumn oCol = _grdLog.Columns.Item(iCol) as SAPbouiCOM.EditTextColumn;
                        oCol.LinkedObjectType = "2";
                        break;
                    case "Data":
                        break;
                    case "Result":
                        break;
                    case "Success":
                        _grdLog.Columns.Item(iCol).Type = BoGridColumnType.gct_CheckBox;
                        break;
                    case "Direction":
                        _grdLog.Columns.Item(iCol).Type = BoGridColumnType.gct_ComboBox;
                        cboCol = (SAPbouiCOM.ComboBoxColumn)_grdLog.Columns.Item(iCol);
                        cboCol.FillValidValues(Src.Resource.Queries.OFTLOG_GET_COMBO_BOX_DIRECTION);
                        break;
                    case "ObjType":
                        _grdLog.Columns.Item(iCol).Visible = false;
                        break;
                }
            }
            _grdLog.AutoResizeColumns();
        }

    }

}
