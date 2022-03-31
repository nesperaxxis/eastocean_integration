using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AXC_EOA_WMSIntegration.Src.Support
{
    public delegate void BeforeFormCloseEventHandler(SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent);
    public delegate void AfterFormLoadEventHandler(SAPbouiCOM.SBOItemEventArg pVal);

    public class FormEvent
    {
        public event BeforeFormCloseEventHandler beforeFormClose;
        public event AfterFormLoadEventHandler afterFormLoad;
        public static System.Collections.Generic.Dictionary<String, SAPbouiCOM.EventForm> registeredFormEvents = new Dictionary<string, SAPbouiCOM.EventForm>();
        private String _FormType = "";
        private String _FormUID = "";
        SAPbouiCOM.EventForm oEventForm;

        public FormEvent(String FormType, String FormUID)
        {
            
            if (!registeredFormEvents.Keys.Contains(FormType))
            {
                oEventForm = Addon.SBO_Application.Forms.GetEventForm(FormType);
                oEventForm.CloseBefore += OnBeforeFormClose;
                oEventForm.LoadAfter += OnAfterFormload;
                registeredFormEvents.Add(FormType, oEventForm);
            }
            else
            {
                oEventForm = registeredFormEvents[FormType];
                oEventForm.CloseBefore -= OnBeforeFormClose;
                oEventForm.CloseBefore += OnBeforeFormClose;
                oEventForm.LoadAfter -= OnAfterFormload;
                oEventForm.LoadAfter += OnAfterFormload;
            }

            _FormType = FormType;
            _FormUID = FormUID;
        }

        void OnAfterFormload(SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.FormUID != _FormUID)
                return;

            AfterFormLoadEventHandler handler = afterFormLoad;
            if (handler != null)
                handler(pVal);
        }

        protected virtual void OnBeforeFormClose(SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            if (pVal.FormUID != _FormUID)
            {
                BubbleEvent = true;
                return;
            }
            BeforeFormCloseEventHandler handler = beforeFormClose;
            if (handler != null)
            {
                handler(pVal, out BubbleEvent);
            }
            else
                BubbleEvent = true;
        }

        public void deRegister()
        {
            oEventForm.CloseBefore -= OnBeforeFormClose;
        }

    }

}
