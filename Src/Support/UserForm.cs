using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AXC_EOA_WMSIntegration.Src.Support
{
    public abstract class UserForm
    {
        protected SAPbouiCOM.Form _oForm;
        protected FormEvent _myFormEvent;

        protected abstract void GetItemReferences();
        protected abstract void InitForm();
        protected abstract void OnBeforeFormClose(SAPbouiCOM.SBOItemEventArg pVal, out bool Bubble);


        public UserForm()
        {
            try
            {
                //Draw the form
                String sFileName = string.Format("{0}.xml", this.GetType().Name);
                if (System.IO.File.Exists(Environment.CurrentDirectory + "\\" + sFileName))
                {
                    System.Xml.XmlDocument oXml = new System.Xml.XmlDocument();
                    oXml.Load(Environment.CurrentDirectory + "\\" + sFileName);
                    String sXml = ModifySize(oXml.InnerXml);
                    Addon.SBO_Application.LoadBatchActions(ref sXml);
                }
                else
                {
                    String ResourceName = string.Format("{0}.Src.Resource.{1}", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, sFileName);

                    String sXml = ModifySize(eCommon.GetXMLResource(ResourceName));
                    Addon.SBO_Application.LoadBatchActions(ref sXml);
                }
                _oForm = Addon.SBO_Application.Forms.ActiveForm;
                if (_oForm.TypeEx.StartsWith("-"))
                {
                    //a UDF form is opened. Close it.
                    String UDFFormUID = _oForm.UniqueID;
                    String ParentFormUID = eCommon.GetParentFormUID(_oForm);
                    _oForm.Close();

                    _oForm = Addon.SBO_Application.Forms.ActiveForm;
                    if (_oForm.UniqueID != ParentFormUID)
                        _oForm = Addon.SBO_Application.Forms.Item(ParentFormUID);
                }
                _oForm.EnableMenu("6913", false);

                RegisterFormEvents();

                GetItemReferences();
                InitForm();
                if (!Addon.oOpenForms.Contains(_oForm.UniqueID))
                    Addon.oOpenForms.Add(_oForm.UniqueID, this);



                _oForm.Visible = true;

                
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("(-10)")) return;   //No Authorization
                Addon.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }


        public UserForm(SAPbouiCOM.Form oForm)
        {
            _oForm = oForm;
            RegisterFormEvents();

            GetItemReferences();
            if (!SBOAddon.oOpenForms.Contains(_oForm.UniqueID))
                SBOAddon.oOpenForms.Add(_oForm.UniqueID, this);
        }

        public UserForm(String FormUID)
        {
            _oForm = Addon.SBO_Application.Forms.Item(FormUID);
            RegisterFormEvents();
            

            GetItemReferences();
            if (!Addon.oOpenForms.Contains(_oForm.UniqueID))
                Addon.oOpenForms.Add(_oForm.UniqueID, this);

        }

        private void RegisterFormEvents()
        {
            _myFormEvent = new FormEvent(_oForm.TypeEx, _oForm.UniqueID);
            _myFormEvent.beforeFormClose += _OnBeforeFormClose;
        }

        private void deRegisterFormEvents()
        {
            _myFormEvent.beforeFormClose -= _OnBeforeFormClose;
            _myFormEvent.deRegister();
            _myFormEvent = null;
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }


        private void _OnBeforeFormClose(SAPbouiCOM.SBOItemEventArg pVal, out bool Bubble)
        {

            Bubble = true;
            try
            {
                OnBeforeFormClose(pVal, out Bubble);
            }
            catch (Exception Ex)
            {
                Addon.SBO_Application.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                Bubble = false;
            }

            if (Bubble)
            {
                try
                {
                    deRegisterFormEvents();
                    for (int i = 0; i < _oForm.DataSources.DataTables.Count; i++)
                    {
                        _oForm.DataSources.DataTables.Item(i).Clear();
                    }

                }
                finally
                {
                    if (Addon.oOpenForms.Contains(_oForm.UniqueID))
                        Addon.oOpenForms.Remove(_oForm.UniqueID);

                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }
            }

        }

        /// <summary>
        /// This will take the XML string and modify the form left/width, top/height to match the current client font size
        /// </summary>
        /// <param name="Original">Original XML String</param>
        /// <returns></returns>
        private static string ModifySize(string Original)
        {

            System.Collections.Specialized.OrderedDictionary oTypes = new System.Collections.Specialized.OrderedDictionary();
            double dFontHeightRatio = Math.Round(Addon.SBO_Application.GetFormItemDefaultHeight(SAPbouiCOM.BoFormSizeableItemTypes.fsit_EDIT) / 14.00, 2);        //Ratio is based on Edit text item. 14.00 is the reference Height that i created the forms in
            double dFontWidthRatio = Math.Round(Addon.SBO_Application.GetFormItemDefaultWidth(SAPbouiCOM.BoFormSizeableItemTypes.fsit_EDIT) / 80.00, 2);                            //Ratio is based on Edit text item. 80.00 is the reference Width that i created the forms in

            oTypes.Add("left", dFontWidthRatio);
            oTypes.Add("width", dFontWidthRatio);
            oTypes.Add("top", dFontHeightRatio);
            oTypes.Add("height", dFontHeightRatio);

            foreach (string Type in oTypes.Keys)
            {
                int i = 0;
                double Ratio = (double)oTypes[Type];
                while (i < Original.Length)
                {
                    i = Original.IndexOf(Type + "=\"", i);
                    if (i > 0)
                    {
                        int iNextApos = Original.IndexOf("\"", i + Type.Length + 2);
                        string sContent = Original.Substring(i + Type.Length + 2, iNextApos - (i + Type.Length + 2));
                        int iContent = 0;
                        if (int.TryParse(sContent, out iContent))
                        {
                            Original = Original.Substring(0, i) + Type + "=\"" + Convert.ToInt16(iContent * Ratio).ToString() + Original.Substring(iNextApos);
                        }
                        i = iNextApos;
                    }
                    else
                    {
                        i = Original.Length;
                    }
                }
            }
            return Original;
        }




    }
}
