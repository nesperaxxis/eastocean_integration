using System;
using System.Collections.Generic;
using System.Text;

namespace AXC_EOA_WMSIntegration
{
    using SAPbobsCOM;
    using SAPbouiCOM;
    using B1WizardBase;
    using System;
    public class SBOAddon_DB:B1Db
    {
        public const string UDO_CODE_SETUP = "AXC_OFTIS";

        public const string SETUP_UDO_RECORD_ID = "SETUP";
        public const string OFTIS_WS_ADDRESS = "U_AXC_WSADR";
        public const string OFTIS_WS_USERNAME = "U_AXC_WSUSN";
        public const string OFTIS_WS_PASSWORD = "U_AXC_WSPWD";
        public const string OFTIS_WS_AUTO_SYNCH_BUSINESS_PARTNERS = "U_AXC_AOCRD";
        public const string OFTIS_WS_AUTO_SYNCH_ITEMS = "U_AXC_AOITM";
        //public const string OFTIS_WS_AUTO_SYNCH_BRANDS = "U_AXC_AOMRC";       --Brands cannot be auto synched, the form is a table list 
        public const string OFTIS_WS_AUTO_SYNCH_ITEM_BARCODES = "U_AXC_AOBCD";
        public const string OFTIS_WS_AUTO_SYNCH_ITEM_GROUPS = "U_AXC_AOITB";
        public const string OFTIS_WS_AUTO_SYNCH_WAREHOUSES = "U_AXC_AOWHS";
        public const string OFTIS_WS_AUTO_SYNCH_BIN_LOCATIONS = "U_AXC_AOBIN";
        public const string OFTIS_WS_AUTO_SYNCH_BILL_OF_MATERIALS = "U_AXC_AOITT";
        public const string OFTIS_WS_AUTO_SYNCH_SALES_ORDERS = "U_AXC_AORDR";
        public const string OFTIS_WS_AUTO_SYNCH_AR_RESERVE_INVOICES = "U_AXC_AOINV";
        public const string OFTIS_WS_AUTO_SYNCH_AR_RETURNS = "U_AXC_AORDN";
        public const string OFTIS_WS_AUTO_SYNCH_AR_CNS = "U_AXC_AORIN";
        public const string OFTIS_WS_AUTO_SYNCH_PURCHASE_ORDERS = "U_AXC_AOPOR";
        public const string OFTIS_WS_AUTO_SYNCH_AP_RETURNS = "U_AXC_AORPD";
        public const string OFTIS_WS_AUTO_SYNCH_AP_CNS = "U_AXC_AORPC";
        public const string OFTIS_WS_AUTO_SYNCH_WORK_ORDERS = "U_AXC_AOWOR";
        public const string OFTIS_WS_AUTO_SYNCH_STOCK_COUNTING = "U_AXC_AOINC";

        public const string FTIS1_ALERT_USERID = "U_AXC_USRID";
        public const string FTIS1_ALERT_USERNAME = "U_AXC_UNAME";
        public const string FTIS1_ALERT_PICK_LIST = "U_AXC_ALPKL";
        public const string FTIS1_ALERT_GRPO = "U_AXC_ALPDN";
        public const string FTIS1_ALERT_INV_TF_REQUEST = "U_AXC_ALWTQ";
        public const string FTIS1_ALERT_INV_TF = "U_AXC_ALWTR";
        public const string FTIS1_ALERT_ISSUE_PROD = "U_AXC_ALGEP";
        public const string FTIS1_ALERT_RECEIPT_PROD = "U_AXC_ALGNP";
        public const string FTIS1_ALERT_STOCK_ADJ_POS = "U_AXC_ALIGN";
        public const string FTIS1_ALERT_STOCK_ADJ_NEG = "U_AXC_ALIGE";
        public const string FTIS1_ALERT_DO_RETURN = "U_AXC_ALRDN";
        public const string FTIS1_ALERT_DELIVERY_ORDER = "U_AXC_ALDLN";
        public const string FTIS1_ALERT_STOCK_POSTING = "U_AXC_ALIQR";

        public const string OFTLG_TABLE_UID = "@AXC_OFTLG";
        public const string OFTLG_USER_ID = "U_AXC_OUSER";
        public const string OFTLG_WS_OBJECT_TYPE = "U_AXC_OBJTP";
        public const string OFTLG_WS_OBJECT_CODE = "U_AXC_OBJCD";
        public const string OFTLG_WS_OBJECT_NAME = "U_AXC_OBJNM";
        public const string OFTLF_WS_DIRECTION = "U_AXC_DRCTN";
        public const string OFTLG_WS_OPERATION = "U_AXC_OPRTN";
        public const string OFTLG_WS_POST_DATA = "U_AXC_PDATA";
        public const string OFTLG_WS_POST_RESULT = "U_AXC_PRSLT";
        public const string OFTLG_WS_POST_SUCCESS = "U_AXC_SCCES";
        public const string OFTLG_WS_EXPORT_TIME_STAMP = "U_AXC_TSTMP";
        public const string OFTLG_WS_EXTERNAL_KEY = "U_AXC_EXTID";

        public const string ODOC_UDF_EXTERNAL_ID = "U_AXC_EXTID";
        public const string DOC1_UDF_EXTERNAL_ID = "U_AXC_EXTID";

        public const string OITM_UDF_PRODUCT_RANKING = "U_PRANK";


        public SBOAddon_DB():base() {
            SAPbobsCOM.Recordset ors = SBOAddon.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;

            try
            {
                //Check for required UDF, UDT and UDOs
                ors.DoQuery("SELECT TOP 1 U_AXC_WSADR, U_AXC_WSUSN, U_AXC_WSPWD, U_AXC_AOCRD, U_AXC_AOITM, U_AXC_AOBCD, U_AXC_AOITB, U_AXC_AOWHS, U_AXC_AORPC, U_AXC_AORPD, U_AXC_AOINV, U_AXC_AORDN, U_AXC_AORIN, U_AXC_AOITT, U_AXC_AOBIN, U_AXC_AOPOR, U_AXC_AORDR, U_AXC_AOWOR, U_AXC_AOINC FROM \"@AXC_OFTIS\"");
                ors.DoQuery("SELECT TOP 1 U_AXC_USRID, U_AXC_UNAME, U_AXC_ALPKL, U_AXC_ALPDN, U_AXC_ALWTQ, U_AXC_ALWTR, U_AXC_ALGEP, U_AXC_ALGNP, U_AXC_ALIGN, U_AXC_ALIGE, U_AXC_ALRDN, U_AXC_ALDLN, U_AXC_ALIQR FROM \"@AXC_FTIS1\""); 
                ors.DoQuery("SELECT TOP 1 U_AXC_OUSER, U_AXC_OBJTP, U_AXC_OBJCD, U_AXC_OBJNM, U_AXC_DRCTN, U_AXC_OPRTN, U_AXC_PDATA, U_AXC_PRSLT, U_AXC_SCCES, U_AXC_TSTMP, U_AXC_EXTID FROM \"@AXC_OFTLG\"");
                ors.DoQuery("SELECT TOP 1 U_AXC_DDLVR, U_AXC_EXTID FROM OPOR");
                ors.DoQuery("SELECT TOP 1 U_AXC_EXTID FROM POR1");
                ors.DoQuery("SELECT TOP 1 U_AXC_SYNCH FROM OCRD");
                ors.DoQuery("SELECT TOP 1 U_AXC_EXTID FROM OPKL");
                ors.DoQuery("SELECT TOP 1 U_AXC_EXTID FROM PKL1");
                ors.DoQuery("SELECT TOP 1 U_AXC_EXTID FROM IQR1");
                ors.DoQuery("SELECT TOP 1 U_PRANK FROM OITM");

                ors.DoQuery("SELECT Count(*) FROM  OUDO WHERE Code in ('AXC_OFTIS')");
                if ((int)ors.Fields.Item(0).Value < 1)
                {
                    throw new Exception("UDO");
                }
            }
            catch 
            {
                //One or more metadata not found. try to recreate them.
                if (ors != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ors);
                    ors = null;
                    GC.Collect();
                }

                Src.Support.Addon.SBO_Application.StatusBar.SetText(SBOAddon.gcAddOnName + " - Setting User Objects.", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning);
                Tables = new B1DbTable[]{
                    new B1DbTable("@AXC_OFTIS", "Integration Setup", BoUTBTableType.bott_MasterData)
                    , new B1DbTable("@AXC_FTIS1", "Integration Setup - Alerts", BoUTBTableType.bott_MasterDataLines)
                    , new B1DbTable("@AXC_OFTLG", "Integration log", BoUTBTableType.bott_NoObjectAutoIncrement)
                };

                Columns = new B1DbColumn[]{
                    new B1DbColumn("OPOR", "AXC_EXTID", "Ext Ref Number", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50, true, new B1DbValidValue[0], -1)
                    , new B1DbColumn("POR1", "AXC_EXTID", "Ext Ref Number", BoFieldTypes.db_Memo, BoFldSubTypes.st_None, 4000, true, new B1DbValidValue[0], -1)
                    , new B1DbColumn("ORDR", "AXC_DDLVR", "Direct Delivery", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("-",""), new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},0)                    
                    , new B1DbColumn("OPKL", "AXC_EXTID", "Ext Ref Number", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50, true, new B1DbValidValue[0], -1)
                    , new B1DbColumn("PKL1", "AXC_EXTID", "Ext Ref Number", BoFieldTypes.db_Memo, BoFldSubTypes.st_None, 4000, true, new B1DbValidValue[0], -1)
                     , new B1DbColumn("OIQR", "AXC_EXTID", "Ext Ref Number", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50, true, new B1DbValidValue[0], -1)
                    , new B1DbColumn("IQR1", "AXC_EXTID", "Ext Ref Number", BoFieldTypes.db_Memo, BoFldSubTypes.st_None, 4000, true, new B1DbValidValue[0], -1)
                    , new B1DbColumn("OITM", "PRANK", "Product Ranking", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50, true, new B1DbValidValue[0], -1)
                    , new B1DbColumn("OCRD", "AXC_SYNCH", "Exclude Integration", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},0)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_WSADR", "API Base Address", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 200, true, new B1DbValidValue[0],-1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_WSUSN", "API UserName", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50, true, new B1DbValidValue[0],-1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_WSPWD", "API Password", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 250, true, new B1DbValidValue[0],-1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AOCRD", "Auto synch BPs", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AOITM", "Auto synch Items", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")}, 1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AOBCD", "Auto synch Barcodes", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AOITB", "Auto synch Item Categories", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AOWHS", "Auto synch Warehouses", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AOBIN", "Auto synch Bin Codes", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AOITT", "Auto synch BOM", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AOWOR", "Auto synch Work Orders", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")}, 1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AORDR", "Auto synch Sales Orders", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AOINV", "Auto synch Reserve Invoice", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AORDN", "Auto synch Sales Returns", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")}, 1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AORIN", "Auto synch AR CNs", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")}, 1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AOPOR", "Auto synch Purchase Orders", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AORPD", "Auto synch Purchase Returns", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")}, 1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AORPC", "Auto synch AP CNs", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_AOINC", "Auto synch Stock Counting", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},1)
                    , new B1DbColumn("@AXC_OFTIS", "AXC_LOGDB", "WS Log DB", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50, true, new B1DbValidValue[0], -1)
                    , new B1DbColumn("@AXC_FTIS1", "AXC_USRID", "User ID", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 200, true, new B1DbValidValue[0],-1)
                    , new B1DbColumn("@AXC_FTIS1", "AXC_UNAME", "User Name", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50, true, new B1DbValidValue[0],-1)
                    , new B1DbColumn("@AXC_FTIS1", "AXC_ALPKL", "Alert PickList", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},0)
                    , new B1DbColumn("@AXC_FTIS1", "AXC_ALPDN", "Alert GRPO", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},0)
                    , new B1DbColumn("@AXC_FTIS1", "AXC_ALWTQ", "Alert Inv.Tr.Rq", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},0)
                    , new B1DbColumn("@AXC_FTIS1", "AXC_ALWTR", "Alert Inv.Tr", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},0)
                    , new B1DbColumn("@AXC_FTIS1", "AXC_ALGEP", "Alert R.Prod", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},0)
                    , new B1DbColumn("@AXC_FTIS1", "AXC_ALGNP", "Alert I.Prod", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},0)
                    , new B1DbColumn("@AXC_FTIS1", "AXC_ALIGN", "Alert Stock+", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},0)
                    , new B1DbColumn("@AXC_FTIS1", "AXC_ALIGE", "Alert Stock-", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},0)
                    , new B1DbColumn("@AXC_FTIS1", "AXC_ALRDN", "Alert DO Return", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},0)
                    , new B1DbColumn("@AXC_FTIS1", "AXC_ALDLN", "Alert Delivery", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},0)
                    , new B1DbColumn("@AXC_FTIS1", "AXC_ALIQR", "Alert Stock Post", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")},0)
                    , new B1DbColumn("@AXC_OFTLG", "AXC_OUSER", "UserCode", BoFieldTypes.db_Numeric, BoFldSubTypes.st_None, 11, true, new B1DbValidValue[0], -1)
                    , new B1DbColumn("@AXC_OFTLG", "AXC_OBJTP", "Object Type", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 25, true, new B1DbValidValue[0], -1)
                    , new B1DbColumn("@AXC_OFTLG", "AXC_OBJCD", "Object Code", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, true, new B1DbValidValue[0], -1)
                    , new B1DbColumn("@AXC_OFTLG", "AXC_OBJNM", "Object Name", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None,254, true, new B1DbValidValue[0], -1)
                    , new B1DbColumn("@AXC_OFTLG", "AXC_DRCTN", "Direction", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("I","Inbound"), new B1DbValidValue("O","Outbound")}, 0)
                    , new B1DbColumn("@AXC_OFTLG", "AXC_OPRTN", "Operation", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, true, new B1DbValidValue[] {new B1DbValidValue("POST","POST"), new B1DbValidValue("PUT","PUT"), new B1DbValidValue("GET","GET"), new B1DbValidValue("DELETE","DELETE"), new B1DbValidValue("PATCH","PATCH")}, 0)
                    , new B1DbColumn("@AXC_OFTLG", "AXC_PDATA", "Post Data", BoFieldTypes.db_Memo, BoFldSubTypes.st_None, 4000, true, new B1DbValidValue[0], -1)
                    , new B1DbColumn("@AXC_OFTLG", "AXC_PRSLT", "Post Result", BoFieldTypes.db_Memo, BoFldSubTypes.st_None, 4000, true, new B1DbValidValue[0], -1)
                    , new B1DbColumn("@AXC_OFTLG", "AXC_SCCES", "Success", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, true, new B1DbValidValue[] {new B1DbValidValue("N","No"), new B1DbValidValue("Y","Yes")}, 0)
                    , new B1DbColumn("@AXC_OFTLG", "AXC_TSTMP", "TimeStamp", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 25, true, new B1DbValidValue[0], -1)
                    , new B1DbColumn("@AXC_OFTLG", "AXC_EXTID", "External ID", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, true, new B1DbValidValue[0], -1)
                    }; 

                Udos = new B1Udo[]{
                    new B1Udo("AXC_OFTIS", "Integration Setup", "AXC_OFTIS", new String[] {"AXC_FTIS1"}, BoUDOObjType.boud_MasterData, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tYES, BoYesNoEnum.tNO, "AAXC_OFTIS", new String[] {"Code"}, new String[] {"Name"})
                };

                try
                {
                    SBOAddon.SBO_Application.MetadataAutoRefresh = false;
                    this.Add(AXC_EOA_WMSIntegration.Src.Support.Addon.oCompany);
                }
                finally
                {
                    SBOAddon.SBO_Application.MetadataAutoRefresh = true;
                }
            }
        }

    }
}
