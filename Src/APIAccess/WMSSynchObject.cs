using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{
    internal abstract class WMSSynchObject
    {
        internal abstract string wmsAPIEndPoint { get; }
        internal abstract string wmsObjectType { get; }
        internal abstract int sapObjectType { get; }
        internal abstract string sapKeyField { get; }
        internal abstract string sapNameField { get; }
        internal abstract string sapKeyVal { get; }

        public WMSSynchObject() { }
        internal virtual string GetJsonObjectPayload()
        {
            return Newtonsoft.Json.JsonConvert.SerializeObject(this);
        }


    }
}
