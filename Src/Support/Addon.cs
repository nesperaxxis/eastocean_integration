using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AXC_EOA_WMSIntegration.Src.Support
{
    public abstract class Addon
    {
        public static SAPbobsCOM.Company oCompany;
        public static SAPbouiCOM.Application SBO_Application;
        private static System.IO.StreamWriter oEventLog = null;

        public static string gcAddOnName = "";
        public static string gcAddonString = "";
        public static String WorkingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\" + gcAddOnName;
        public static System.Collections.Hashtable Forms = new System.Collections.Hashtable();
        public static System.Collections.Specialized.OrderedDictionary oOpenForms = new System.Collections.Specialized.OrderedDictionary();
        public static System.Collections.Specialized.OrderedDictionary oFormEvents = new System.Collections.Specialized.OrderedDictionary();
        public static System.Collections.Hashtable oRegisteredFormEvents = new System.Collections.Hashtable();
        public Boolean Connected = false;
        
        public const string Filler = "      ";
        public static int iPriceAccuracy = 2;
        public static int iQtyAccuracy = 0;
        public static int iAmountAccuracy = 2;
        public static string LocalCurr = "";
        public static string SystemCurr = "";
        public static string TempField = "";
        public static bool isSegmentedAcc = false;
        public static string sAcctSeparator = "";

        protected abstract void OnMenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool Bubble);
        protected abstract void OnAppEvents(SAPbouiCOM.BoAppEventTypes EventType);

        /// <summary>
        /// Add menu tree here in case you need a drop down menu.
        /// </summary>
        protected virtual void CreateMenuTree() { }
        protected virtual void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        /// <summary>
        /// Initializing the base Addon
        /// 
        /// </summary>
        /// <param name="Args">Contains the connection string to the running SBO Application</param>
        /// <param name="AddonCode">This code will be used to identify the current Addon, Main Menu UID, User Authorization</param>
        /// <param name="AddonName">This name will shows up as the main menu name.</param>
        public Addon(String[] Args, String AddonCode, String AddonName )
        {
            gcAddOnName = AddonCode;
            gcAddonString = AddonName;
            WorkingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\" + gcAddOnName;

            try
            {
                SAPbouiCOM.SboGuiApi oGUI = new SAPbouiCOM.SboGuiApi();
                if (Args.Length == 0)
                    oGUI.Connect("0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056");
                else
                    oGUI.Connect(Args[0]);


                SetApplication(oGUI.GetApplication(-1), gcAddOnName, true);
                SBOAddon_DB addOnDb = new SBOAddon_DB();
                
                //Application forms
                Forms = CollectFormsAttribute();
                //Register Events
                RegisterAppEvents();
                RegisterFormEvents();

                //Register currently opened forms - initialized opened forms so it is ready to use.
                RegisterForms();

                //Create Authorization
                AddAuthorizationTree();

                //Add the menus
                CreateMenuFromFormAttribute();

            }
            catch (Exception Ex)
            {
                throw new Exception("ERROR - Connection failed: " + Ex.Message); ;
            }
        }

        private System.Collections.Hashtable CollectFormsAttribute()
        {
            System.Collections.Hashtable oTable = new System.Collections.Hashtable();
            string NameSpace = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            Type FormAttr = Type.GetType(string.Format("{0}.FormAttribute", NameSpace));
            foreach (System.Reflection.Assembly asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (asm.FullName.StartsWith("mscorlib"))
                    continue;
                if (asm.FullName.StartsWith("Interop"))
                    continue;
                if (asm.FullName.StartsWith("System"))
                    continue;
                if (asm.FullName.StartsWith("Microsoft"))
                    continue;

                foreach (Type type in asm.GetTypes())
                {
                    foreach (System.Attribute Attr in type.GetCustomAttributes(FormAttr, false))
                    {
                        FormAttribute frmAttr = (FormAttribute)Attr;
                        frmAttr.TypeName = type.FullName;
                        if (!oTable.ContainsKey(frmAttr.FormType))
                            oTable.Add(frmAttr.FormType, frmAttr);
                        else
                            SBO_Application.MessageBox(string.Format("The form type {0} can not be registered twice", frmAttr.FormType));
                    }

                }
            }

            return oTable;

        }

        private System.Collections.Hashtable CollectAuthorizationAttribute()
        {
            System.Collections.Hashtable oTable = new System.Collections.Hashtable();
            string NameSpace = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            Type AuthAttrType = Type.GetType(string.Format("{0}.AuthorizationAttribute", NameSpace));

            foreach (System.Reflection.Assembly asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (asm.FullName.StartsWith("mscorlib"))
                    continue;
                if (asm.FullName.StartsWith("Interop"))
                    continue;
                if (asm.FullName.StartsWith("System"))
                    continue;
                if (asm.FullName.StartsWith("Microsoft"))
                    continue;

                foreach (Type type in asm.GetTypes())
                {
                    foreach (System.Attribute Attr in type.GetCustomAttributes(AuthAttrType, false))
                    {
                        AuthorizationAttribute AuthAttr = (AuthorizationAttribute)Attr;
                        oTable.Add(AuthAttr.FormType, AuthAttr);
                    }
                }
            }

            return oTable;

        }

        private void AddAuthorizationTree()
        {
            try
            {
                SAPbobsCOM.UserPermissionTree oUserPer = (SAPbobsCOM.UserPermissionTree)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree);
                int lErr = 0;

                if (oUserPer.GetByKey(SBOAddon.gcAddOnName) == false)
                {
                    oUserPer.PermissionID = SBOAddon.gcAddOnName;
                    oUserPer.Name = string.Format("Addon : {0}", SBOAddon.gcAddOnName);
                    oUserPer.Options = SAPbobsCOM.BoUPTOptions.bou_FullReadNone;
                    lErr = oUserPer.Add();
                    if (lErr != 0)
                        throw new Exception(oCompany.GetLastErrorDescription());
                }

                System.Collections.Hashtable AuthTable = CollectAuthorizationAttribute();
                foreach (string FormType in AuthTable.Keys)
                {
                    AuthorizationAttribute AuthAttrib = AuthTable[FormType] as AuthorizationAttribute;
                    oUserPer = (SAPbobsCOM.UserPermissionTree)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree);
                    if (oUserPer.GetByKey(AuthAttrib.FormType) == false)
                    {
                        oUserPer.PermissionID = AuthAttrib.FormType;
                        oUserPer.Name = AuthAttrib.Name;
                        if (AuthAttrib.ParentID == "")
                            oUserPer.ParentID = SBOAddon.gcAddOnName;
                        else
                            oUserPer.ParentID = AuthAttrib.ParentID;
                        oUserPer.Options = AuthAttrib.Options;
                        oUserPer.UserPermissionForms.FormType = AuthAttrib.FormType;
                        lErr = oUserPer.Add();
                        if (lErr != 0)
                            throw new Exception(oCompany.GetLastErrorDescription());
                    }
                }



            }
            catch (Exception ex)
            {
                oEventLog.WriteLine(DateTime.Now + Filler + "Unable to create Authorization for " + SBOAddon.gcAddOnName + " Module. " + ex.Message);
            }
        }

        public void RegisterAppEvents()
        {
            SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(OnMenuEvent);
            SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(OnAppEvents);
            SBO_Application.ItemEvent += SBO_Application_ItemEvent;
        }

        /// <summary>
        /// Get the form events based on the Form Event attribute declared on the methods in each of the class
        /// </summary>
        private void RegisterFormEvents()
        {
            string NameSpace = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            Type FormEventAttrType = Type.GetType(string.Format("{0}.FormEventAttribute", NameSpace));

            foreach (System.Reflection.Assembly asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (asm.FullName.StartsWith("mscorlib"))
                    continue;
                if (asm.FullName.StartsWith("Interop"))
                    continue;
                if (asm.FullName.StartsWith("System"))
                    continue;
                if (asm.FullName.StartsWith("Microsoft"))
                    continue;

                foreach (Type type in asm.GetTypes())
                {

                    Type FormAttr = Type.GetType(string.Format("{0}.FormAttribute", NameSpace));
                    FormAttribute frmAttr = null;
                    foreach (System.Attribute Attr in type.GetCustomAttributes(FormAttr, false))
                    {
                        frmAttr = (FormAttribute)Attr;
                    }
                    //Get the methods attribute
                    foreach (System.Reflection.MethodInfo method in type.GetMethods())
                    {
                        foreach (System.Attribute Attr in method.GetCustomAttributes(FormEventAttrType, false))
                        {
                            SAPbouiCOM.EventForm oEvent = null;
                            FormEventAttribute frmEventAttr = (FormEventAttribute)Attr;
                            String sKey = string.Format("{0}_{1}", frmAttr.FormType, frmEventAttr.oEventType.ToString());
                            if (!SBOAddon.oFormEvents.Contains(frmAttr.FormType))
                            {
                                oEvent = SBO_Application.Forms.GetEventForm(frmAttr.FormType);
                                SBOAddon.oFormEvents.Add(frmAttr.FormType, oEvent);
                            }
                            else
                            {
                                oEvent = (SAPbouiCOM.EventForm)SBOAddon.oFormEvents[frmAttr.FormType];
                            }

                            if (SBOAddon.oRegisteredFormEvents.Contains(sKey))
                                throw new Exception(string.Format("The form event method type [{0}] can not be registered twice", sKey));
                            else
                                SBOAddon.oRegisteredFormEvents.Add(sKey, "");

                            Type EventClass = oEvent.GetType();
                            System.Reflection.EventInfo oInfo = EventClass.GetEvent(frmEventAttr.oEventType.ToString());
                            if (oInfo == null)
                            {
                                throw new Exception(string.Format("Invalid method info name. [{0}]", frmEventAttr.oEventType.ToString()));
                            }
                            Delegate d = Delegate.CreateDelegate(oInfo.EventHandlerType, method);

                            oInfo.AddEventHandler(oEvent, d);

                        }

                    }
                }
            }

        }


        private void RegisterForms()
        {
            for (int i = 0; i < SBO_Application.Forms.Count; i++)
            {
                if (!oOpenForms.Contains(SBO_Application.Forms.Item(i).UniqueID))
                {
                    FormAttribute oAttrib = Forms[SBO_Application.Forms.Item(i).TypeEx] as FormAttribute;
                    if (oAttrib != null)
                    {
                        try
                        {
                            //Execute the constructor
                            System.Reflection.Assembly asm = System.Reflection.Assembly.GetExecutingAssembly();
                            Type oType = asm.GetType(oAttrib.TypeName);
                            System.Reflection.ConstructorInfo ctor = oType.GetConstructor(new Type[1] { typeof(String) });
                            if (ctor != null)
                            {
                                object oForm = ctor.Invoke(new Object[1] { SBO_Application.Forms.Item(i).UniqueID });
                            }
                            else
                                throw new Exception("No constructor which accepts the formUID found for form type - " + oAttrib.FormType);
                        }
                        catch (Exception ex)
                        {
                            SBO_Application.MessageBox(ex.Message);
                        }
                    }
                }
            }
        }

        private void CreateMenuFromFormAttribute()
        {
            //--------------- remove and load menus -----------
            if (SBO_Application.Menus.Exists(SBOAddon.gcAddOnName)) SBO_Application.Menus.RemoveEx(SBOAddon.gcAddOnName);
            SBO_Application.Menus.Item("43520").SubMenus.Add(SBOAddon.gcAddOnName, gcAddonString, SAPbouiCOM.BoMenuType.mt_POPUP, 99);

            //Call the custom user menu tree if they specify
            CreateMenuTree();

            System.Data.DataTable dtMenus = new System.Data.DataTable();
            dtMenus.Columns.Add("Idx", typeof(int));
            dtMenus.Columns.Add("Code", typeof(string));
            dtMenus.Columns.Add("Attrib", typeof(FormAttribute));

            foreach (string Key in Forms.Keys)
            {
                FormAttribute oAttr = (FormAttribute)Forms[Key];
                if (oAttr.HasMenu)
                {
                    dtMenus.Rows.Add(oAttr.Position, oAttr.FormType, oAttr);
                }
            }

            using (System.Data.DataView view = new System.Data.DataView(dtMenus, "", "Idx ASC, Code ASC", System.Data.DataViewRowState.CurrentRows))
            {
                foreach (System.Data.DataRowView row in view)
                {
                    FormAttribute oAttr = (FormAttribute)row["Attrib"];
                    if (SBO_Application.Menus.Exists(oAttr.FormType)) SBO_Application.Menus.RemoveEx(oAttr.FormType);
                    if (oAttr.ParentMenu == "")
                        SBO_Application.Menus.Item(Addon.gcAddOnName).SubMenus.Add(oAttr.FormType, oAttr.MenuName, SAPbouiCOM.BoMenuType.mt_STRING, oAttr.Position);
                    else
                        SBO_Application.Menus.Item(oAttr.ParentMenu).SubMenus.Add(oAttr.FormType, oAttr.MenuName, SAPbouiCOM.BoMenuType.mt_STRING, oAttr.Position);
                }
            }


            try
            {
                SBO_Application.Menus.Item(SBOAddon.gcAddOnName).Image = Environment.CurrentDirectory + "\\icon.png";
            }
            catch { }
        }

        private void SetApplication(SAPbouiCOM.Application oApp, String AddOnName, bool Multiple)
        {
            SBO_Application = oApp;
            oCompany = new SAPbobsCOM.Company();
            string cookie = oCompany.GetContextCookie();
            string connInfo = SBO_Application.Company.GetConnectionContext(cookie);
            int retCode = oCompany.SetSboLoginContext(connInfo);
            if (retCode == 0)
            {
                if (Multiple)
                    oCompany = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
                else
                {
                    oCompany.Connect();
                }
                //Create a Log File
                CreateLogFile();

                // CHECK AXC ADDON LICENCE
#if !DEBUG
                if (!IsDesignTime())
                {
                    if (!CheckLicenceEX(AddOnName))
                    {
                        throw new System.Runtime.InteropServices.COMException("No Addon Licence.");
                    }
                }
#endif

                //Decimal Accuracies
                SAPbobsCOM.CompanyService oCompanyInfo = oCompany.GetCompanyService();
                iPriceAccuracy = oCompanyInfo.GetAdminInfo().PriceAccuracy;
                iQtyAccuracy = oCompanyInfo.GetAdminInfo().AccuracyofQuantities;
                iAmountAccuracy = oCompanyInfo.GetAdminInfo().TotalsAccuracy;


                //Initialize some values
                SAPbobsCOM.SBObob Sbob = (SAPbobsCOM.SBObob)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                SAPbobsCOM.Recordset oRS = Sbob.GetLocalCurrency();
                LocalCurr = (string)oRS.Fields.Item(0).Value;
                oRS = Sbob.GetSystemCurrency();
                SystemCurr = (string)oRS.Fields.Item(0).Value;
                oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery("SELECT COUNT(*) FROM OASG");
                if ((int)oRS.Fields.Item(0).Value > 0)
                {
                    isSegmentedAcc = true;
                    sAcctSeparator = oCompanyInfo.GetAdminInfo().AccountSegmentsSeparator;
                }
                else
                {
                    isSegmentedAcc = false;
                }

                eCommon.ReleaseComObject((object)oCompanyInfo);
                eCommon.ReleaseComObject((object)oRS);
                eCommon.ReleaseComObject((object)Sbob);


            }
            else
            {
                throw new Exception(oCompany.GetLastErrorDescription());
            }
        }

        /// <summary>
        /// this method is for init company wihtout UI connection
        /// </summary>
        /// <param name="ServerType"></param>
        /// <param name="Server"></param>
        /// <param name="LicServer"></param>
        /// <param name="CompanyDB"></param>
        /// <param name="UserName"></param>
        /// <param name="Password"></param>
        public static void InitCompany(SAPbobsCOM.BoDataServerTypes ServerType, String Server, String LicServer, String CompanyDB, String UserName, String Password)
        {
            //Create a Log File
            CreateLogFile();

            oCompany = new SAPbobsCOM.Company();
            oCompany.DbServerType = ServerType;
            oCompany.Server = Server;
            oCompany.LicenseServer = LicServer;
            oCompany.CompanyDB = CompanyDB;
            oCompany.UserName = UserName;
            oCompany.Password = Password;

            oCompany.Connect();
            if (!oCompany.Connected)
                throw new Exception(oCompany.GetLastErrorDescription());

        }

        private static bool CheckLicenceEX(string AddOnName)
        {
            try
            {
                AXC_Licence.AXC_Licence oLic = new AXC_Licence.AXC_Licence(SBO_Application, oCompany, AddOnName, new Byte[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 10, 73, 1, 5, 75, 1, 8 });
                if (!oLic.IsValid)
                {
                    SBO_Application.StatusBar.SetText("Could not start addon. " + oLic.LastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oCompany.Disconnect();
                    return false;
                }
                else
                {
                    if (oLic.DaysToExpiry < 10)
                    {
                        if (oLic.DaysToExpiry > 0)
                        {
                            SBO_Application.MessageBox("Your add on " + AddOnName + " will expire in " + oLic.DaysToExpiry + " days. Please contact support for the license.", 1, "OK", null, null);
                        }
                        else
                        {
                            SBO_Application.MessageBox("Your add on " + AddOnName + " expires today. Please contact support for the license.", 1, "OK", null, null);
                        }
                    }
                    return true;
                }
            }
            catch
            {
                Addon.SBO_Application.StatusBar.SetText("Could not start addon. No licence found.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                Addon.oCompany.Disconnect();
                System.Windows.Forms.Application.Exit();
                return false;
            }
        }

        private static bool IsDesignTime()
        {
            try
            {
                System.Diagnostics.Process cProcess = System.Diagnostics.Process.GetCurrentProcess();
                if (cProcess.ProcessName.Contains(".vshost")) //This part is for DesignTime Mode
                    return true;
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }


        private static void CreateLogFile()
        {
            try
            {
                //Logging part
                string sLogDir = WorkingDirectory + "\\EventLog";
                if (!System.IO.Directory.Exists(sLogDir)) System.IO.Directory.CreateDirectory(sLogDir);
                string sLogFN = sLogDir + "\\EventLog" + System.DateTime.Now.ToString("yyyyMM") + ".txt";

                int iCount = 1;
                while (true)
                {
                    try
                    {
                        //oEventLog = System.IO.File.AppendText(sLogFN)
                        if (iCount == 1)
                            InitializeEventLog(sLogFN);
                        else
                        {
                            sLogFN = sLogFN.Replace(".txt", "");
                            InitializeEventLog(sLogFN + "(" + iCount + ").txt");
                        }
                        break;
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.ToLower().Contains("because it is being used by another process"))
                        {
                            iCount += 1;
                            if (iCount > 100)
                            {
                                SBO_Application.StatusBar.SetText(SBOAddon.gcAddOnName + " Unable to start logging. Starting without logging.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                break;
                            }
                        }
                        else
                        {
                            SBO_Application.StatusBar.SetText(SBOAddon.gcAddOnName + " Could not start logging." + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            break;
                        }
                    }
                }

                if (oEventLog != null)
                {
                    oEventLog.AutoFlush = true;

                    oEventLog.WriteLine("________________________________________ NEW SESSION ________________________________________");
                    if (iCount > 1) oEventLog.Write(DateTime.Now.ToString("yyyyMMdd HH:mm:ss") + "  " + iCount + " concurrent sessions detected");
                    oEventLog.WriteLine(DateTime.Now.ToString("yyyyMMdd HH:mm:ss") + "  Start Event Log");
                    oEventLog.WriteLine(DateTime.Now.ToString("yyyyMMdd HH:mm:ss") + "  Initiating Add On");
                }
            }
            catch { }

        }

        private static void InitializeEventLog(string sLogFN)
        {
            oEventLog = System.IO.File.AppendText(sLogFN);
        }

        public static void WriteEventLog(string Log)
        {
            if (oEventLog != null)
            {
                oEventLog.WriteLine(DateTime.Now + "  " + Log);
            }
        }

    }
}
