using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Security;

namespace AXC_EOA_WMSIntegration.Src.APIAccess
{    
    public class message
    {
        //TODO: Implement the actual return message from Simplr WMS 

        //public string resultMessage;
        //public bool success;
        //Updated:04/08/2021
        //Actual Smplr Response
        //{
        //     "Response": {
        //         "Status": "Created",
        //         "StatusCode": "201",
        //         "StatusMessage": "Customer created successfully",
        //         "IDocNo": "CLA001",
        //         "TotalIDoc": "1",
        //         "TotalRecord": "1"
        //     }
        // }
        public JObject Response;               
    }

    public class WMSResponse
    {
        public string Status;
        public string StatusCode;
        public string StatusMessage;
        public string IDocNo;
        public string TotalIDoc;
        public string TotalRecord;
    }

    public class APIServiceAccess
    {
        private static string fPath, fFile;
        public enum Operation { POST, PUT, GET, PATCH, DELETE }
        //public enum PostObjectTypes { Vendor, Customer, Item, ItemGroup, Brand, Barcode, BillOfMaterial, Warehouse, BinLocation,
        //    Salesorder, ReserveInvoice, ARReturn, ARCreditNotes, PurchaseOrder, APReturn, APCreditNotes,
        //    WorkOrder }

        public static Dictionary<string, Type> SynchObjectInfoCollection = new Dictionary<string, Type>
        {
            {APCNotes._objectType,  typeof(APCNotes) } ,
            {APReturns._objectType,  typeof(APReturns) } ,  
            {ARCNotes._objectType,  typeof(ARCNotes) } ,
            {ARResInvoices._objectType,  typeof(ARResInvoices) } ,
            {ARReturns._objectType,  typeof(ARReturns) } ,
            {Barcodes._objectType,  typeof(Barcodes) } ,
            {BillOfMaterials._objectType,  typeof(BillOfMaterials) } ,
            {BinLocations._objectType,  typeof(BinLocations) } ,
            {Brands._objectType,  typeof(Brands) } ,
            {Customer._objectType,  typeof(Customer) } ,
            {Vendor._objectType,  typeof(Vendor) } ,
            {ItemCategories._objectType,  typeof(ItemCategories) } ,
            {Items._objectType,  typeof(Items) } ,
            {SalesOrders._objectType,  typeof(SalesOrders) } ,
            {PurchaseOrders._objectType,  typeof(PurchaseOrders) } ,
            {Warehouses._objectType,  typeof(Warehouses) } ,
            {WorkOrders._objectType,  typeof(WorkOrders) } ,
            {ItemCount._objectType,  typeof(ItemCount) } ,
        };

        public static void newWsService()
        {
            //_wsService = null;
        }

        public static bool SynchObject<T>(object code, string objName, out string result)
        {
            bool success = false;
            result = "Not implemented";
            string payLoad = "";
            string mainBracket = "";
            string fullpayLoad = "";
            string wmsAPIEndPoint = "";
            string sKey = String.Empty;
            string sKeyVal = String.Empty;
            string sObj = String.Empty;
            GetSynchObjectType<T>(out var objectType);
            
            try
            {
                //T obj = (T)Convert.ChangeType(Activator.CreateInstance(typeof(T), new object[] { code }), typeof(T))  ; 
                WMSSynchObject obj = (WMSSynchObject)Activator.CreateInstance(typeof(T), new object[] { code });
                payLoad = obj.GetJsonObjectPayload();
                wmsAPIEndPoint = obj.wmsAPIEndPoint;
                sObj = obj.wmsObjectType;
                sKey = obj.sapKeyField;
                sKeyVal = obj.sapKeyVal;
                
                if(wmsAPIEndPoint == "SalesOrder" || wmsAPIEndPoint == "ReserveInvoice") { mainBracket = "OrderDetails"; }
                else if (wmsAPIEndPoint == "PurchaseOrder") { mainBracket = "PODetails"; }
                else if (wmsAPIEndPoint == "ProductionOrder") { mainBracket = "ProductionOrderDetails"; }
                else if (wmsAPIEndPoint == "APGoodsReturn" || wmsAPIEndPoint == "APCreditNote" || wmsAPIEndPoint == "DOReturn" || wmsAPIEndPoint == "ARCN" ) { mainBracket = "CreditNoteDetails"; }
                else { mainBracket = wmsAPIEndPoint.Equals("WarehouseBin") ? wmsAPIEndPoint.Replace("Warehouse", "") : wmsAPIEndPoint; }
                fullpayLoad = "{\"" + mainBracket + "\":[" + payLoad + "]}";
                
                int iteration = 0;
                success = PostWebAPI(wmsAPIEndPoint, fullpayLoad, out result, ref iteration);
            }
            catch (Exception wsExc)
            {
                success = false;
                result = wsExc.Message;
            }
            finally
            {
                axcOFTLG.GenerateRecord(objectType, code.ToString(), objName, "", Operation.POST, fullpayLoad, result, success);
            }

            fPath = @"D:\WMS\JSON Files\" + mainBracket + @"\";
            fFile = sObj + "_" + sKey + "_" + sKeyVal + "_" + (!success ? "failed":"success") +  ".json";
            WriteToFile(fullpayLoad);

            return success;
        }

        public static bool PostWebAPI(string wmsAPIEndPoint, string payLoad, out string result,ref int iteration)
        {
            bool success = false;
            result = "Not implemented";
            axcFTSetup setup = new axcFTSetup();

            try
            {
                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                string httpAddress = setup.Values[SBOAddon_DB.OFTIS_WS_ADDRESS].ToString();
                string userName = setup.Values[SBOAddon_DB.OFTIS_WS_USERNAME].ToString();
                string password = setup.Values[SBOAddon_DB.OFTIS_WS_PASSWORD].ToString();
                if (string.IsNullOrEmpty(httpAddress))
                    throw new Exception("WMS API Base Address must be provided. Check the setup.");


                //TODO: implement the Simplr URL & authentication here
                string uri = String.Format("{0}/api/{1}", httpAddress, wmsAPIEndPoint);
                var content = new StringContent(payLoad, Encoding.UTF8, "application/json");

                var handler = new HttpClientHandler()
                {
                    Proxy = HttpWebRequest.GetSystemWebProxy()
                };
                var client = new HttpClient(handler);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                var byteArray = Encoding.ASCII.GetBytes(String.Format("{0}:{1}", userName, password));
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                //client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("oAuth2", Convert.ToBase64String(byteArray));

                var resulto = client.PostAsync(uri, content).Result;
                if (resulto.IsSuccessStatusCode)
                {
                    string resultMessage = resulto.Content.ReadAsStringAsync().Result;
                    var resultM = Newtonsoft.Json.JsonConvert.DeserializeObject<message>(resultMessage);
                    var WMSmsg = Newtonsoft.Json.JsonConvert.DeserializeObject<List<WMSResponse>>("[" + resultM.Response + "]");
                    foreach (WMSResponse msg in WMSmsg)
                    {
                        success = (msg.Status == "Created") ? true : false;
                        result = msg.StatusMessage;
                    }                    
                }
                else
                {
                    result = resulto.ReasonPhrase;
                    switch (resulto.StatusCode)
                    {
                        case HttpStatusCode.Unauthorized:
                            result = "Access Denied! Kindly provide a valid credential";
                            break;
                        case HttpStatusCode.Forbidden:
                            result = "Your request did not include an Authentication!";
                            break;
                        case HttpStatusCode.BadRequest:
                            result = ": Invalid format";
                            break;
                        case HttpStatusCode.InternalServerError:
                            //Retry
                            result = resulto.StatusCode.ToString();
                            while(iteration<=2)
                            {
                                iteration +=1;
                                System.Diagnostics.Debug.WriteLine("Internal Server Error iteration " + iteration.ToString());
                                success = PostWebAPI(wmsAPIEndPoint, payLoad, out result,ref iteration);
                                if (result != HttpStatusCode.InternalServerError.ToString())
                                    break;
                            }
                            break;
                        default:
                            if (result == "")
                                result = Enum.GetName(typeof(HttpStatusCode), resulto.StatusCode);
                            break;

                    }
                }
            }
            catch (OperationCanceledException oce)
            {
                // Tell the user that the request timed our or you cancelled a CancellationToken
                success = false;
                result = oce.Message;
            }
            catch (Exception Ex)
            {
                success = false;
                if (Ex.InnerException != null)
                {
                    result = Ex.InnerException.Message;
                    if (Ex.InnerException.InnerException != null)
                        result += " " + Ex.InnerException.InnerException.Message;

                }
                else
                    result = Ex.Message;
            }

            if (result.Length > 254)
                result = result.Substring(0, 254);


            return success;
        }

        private static void GetSynchObjectType<T>(out string objectType)
        {
            //Dictionary<Type, Tuple<string, PostObjectTypes>> types = new Dictionary<Type, Tuple<string, PostObjectTypes>>();
            //types.Add(typeof(APCNotes), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_AP_CN, PostObjectTypes.APCreditNotes));
            //types.Add(typeof(APReturns), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_AP_RETURN, PostObjectTypes.APReturn));
            //types.Add(typeof(ARCNotes), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_AR_CN, PostObjectTypes.ARCreditNotes));
            //types.Add(typeof(ARResInvoices), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_RESERVE_INVOICE, PostObjectTypes.ReserveInvoice));
            //types.Add(typeof(ARReturns), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_AR_RETURNS, PostObjectTypes.ARReturn));
            //types.Add(typeof(Barcodes), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_BAR_CODE, PostObjectTypes.Barcode));
            //types.Add(typeof(BillOfMaterials), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_BOM, PostObjectTypes.BillOfMaterial));
            //types.Add(typeof(BinLocations), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_BIN, PostObjectTypes.BinLocation));
            //types.Add(typeof(Brands), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_BRAND, PostObjectTypes.Brand));
            ////types.Add(typeof(BusinessPartners), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_VENDOR, PostObjectTypes.Vendor));
            //types.Add(typeof(ItemCategories), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_ITEM_CATEGORY, PostObjectTypes.ItemGroup));
            //types.Add(typeof(Items), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_ITEM, PostObjectTypes.Item));
            //types.Add(typeof(PurchaseOrders), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_PURCHASE_ORDER, PostObjectTypes.PurchaseOrder));
            //types.Add(typeof(SalesOrders), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_SALES_ORDER, PostObjectTypes.Salesorder));
            //types.Add(typeof(Warehouses), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_WAREHOUSE, PostObjectTypes.Warehouse));
            //types.Add(typeof(WorkOrders), new Tuple<string, PostObjectTypes>(SBOAddon.SYNCH_O_OBJECT_WORK_ORDER, PostObjectTypes.WorkOrder));


            objectType = SynchObjectInfoCollection.Where(x => x.Value == typeof(T)).Select(y=>y.Key).FirstOrDefault();

        }

        private static Type GetSynchObjectClass(string objectType)
        {
            Type type = SynchObjectInfoCollection[objectType];
            return type;
        }

        public static void WriteToFile(string Message)
        {
            if(!Directory.Exists(fPath))
            {
                Directory.CreateDirectory(fPath);
            }

            if (!File.Exists(fPath + fFile))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(fPath + fFile))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(fPath + fFile))
                {
                    sw.WriteLine(Message);
                }
            }
        }


    }
}
