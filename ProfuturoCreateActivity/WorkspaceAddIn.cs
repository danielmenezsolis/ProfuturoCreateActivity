using System;
using System.AddIn;
using System.Collections.Generic;
using System.Drawing;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Windows.Forms;
using ProfuturoCreateActivity.POCRNService;
using RestSharp;
using RestSharp.Authenticators;
using RightNow.AddIns.AddInViews;


namespace ProfuturoCreateActivity
{
    public class WorkspaceAddIn : Panel, IWorkspaceComponent2
    {

        private string Observaciones { get; set; }
        private string customerName { get; set; }
        private string ApellidoPaterno { get; set; }
        private string ApellidoMaterno { get; set; }
        private string streetAddress { get; set; }
        private string city { get; set; }
        private string postalCode { get; set; }
        private string stateProvince { get; set; }
        private string NSS { get; set; }
        private string Meses { get; set; }
        private string Régimen { get; set; }
        private string CURP { get; set; }
        private string Empresa { get; set; }
        private string Contacto { get; set; }
        private string PhoneCel { get; set; }
        private string PhoneHome { get; set; }
        private string Correo { get; set; }



        private RightNowSyncPortClient clientRN { get; set; }
        private IRecordContext recordContext { get; set; }
        private IGlobalContext globalContext { get; set; }
        private IIncident Incident { get; set; }
        private int IncidentID { get; set; }
        public IGenericObject SWorkOrder;

        public WorkspaceAddIn(bool inDesignMode, IRecordContext recordContext, IGlobalContext globalContext)
        {
            try
            {
                if (!inDesignMode)
                {
                    this.recordContext = recordContext;
                    this.globalContext = globalContext;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace);
            }
        }


        public Control GetControl()
        {
            return this;
        }

        public bool ReadOnly { get; set; }


        public void RuleActionInvoked(string ActionName)
        {
            if (ActionName == "CrearActividad")
            {
                if (Init())
                {
                    SWorkOrder = recordContext.GetWorkspaceRecord("TOA$Work_Order") as IGenericObject;
                    if (SWorkOrder != null)
                    {
                        ObtenerInfoWorkOrder(SWorkOrder.Id);
                        CrearActividad();
                      
                    }
                    else
                    {
                        MessageBox.Show("El objeto es Nulo");
                    }
                }
                else
                {
                    MessageBox.Show("No hay Init");
                }
            }
        }

        public void CrearActividad()
        {
            try
            {
                var client = new RestClient("https://demo-usdc6.etadirect.com");
                client.Authenticator = new HttpBasicAuthenticator("test1234@ofsc-57130f.test", "cbc494cad1463968e025e119c018c1049ab242bd92e9dc2467f2c2d95c096e87");

                var request = new RestRequest("/ofsc-57130f.test/rest/ofscCore/v1/activities/", Method.POST)
                {
                    RequestFormat = DataFormat.Json
                };
                var body = "{";
                body += "\"apptNumber\":\"1001\",";
                body += "\"resourceId\":\"z13m\",";
                body += "\"date\":\"2019-08-13\",";
                body += "\"timeSlot\":\"14:00 - 14:30\",";
                body += "\"activityType\":\"03\",";
                body += "\"customerName\":\"TEST\",";
                body += "\"Apellido Paterno\":\"AP\",";
                body += "\"Apellido Materno\":\"AM\",";
                body += "\"streetAddress\":\"Hamburgo 206\",";
                body += "\"city\":\"Juárez\",";
                body += "\"postalCode\":\"06600\",";
                body += "\"stateProvince\":\"CDMX\",";
                body += "\"NSS\":\"62636377374\",";
                body += "\"Observaciones\":\"carta para viernes 9 a las 11am  insurgentes banorte xxi\",";
                body += "\"Meses en su afore actual\":\"12\",";
                body += "\"Tipo de cartera\":\"Cedidos\",";
                body += "\"Afore\":\"Invercap\",";
                body += "\"Régimen\":\"Ley 73\",";
                body += "\"Estatus del trámite\":\"CIT\",";
                body += "\"Curp\":\"12345674\",";
                body += "\"Nombre empresa\":\"EmpresaEsa\",";
                body += "\"customer_number\":\"CustoNameEsa\",";
                body += "\"cphone\":\"7225251549\",";
                body += "\"ccell\":\"7225251549\",";
                body += "\"cemail\":\"daniel_menez@hotmail.com\"";


                //body += "\"Costo\":\"" + costo + "\"";



                body += "}";
                globalContext.LogMessage("Crear Actividad: \n " + body);
                request.AddParameter("application/json", body, ParameterType.RequestBody);



                // execute the request
                IRestResponse response = client.Execute(request);
                var content = response.Content; // raw content as string
                if (content == "")
                {

                }
                else
                {
                    MessageBox.Show(response.Content);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("UpdatePaxPrice:" + ex.Message + " Det: " + ex.StackTrace);
            }

        }
        private void ObtenerInfoWorkOrder(int WorkId)
        {
            try
            {
                ClientInfoHeader clientInfoHeader = new ClientInfoHeader();
                APIAccessRequestHeader aPIAccessRequest = new APIAccessRequestHeader();
                clientInfoHeader.AppID = "Query Example";
                String queryString = "SELECT case_note,contact.LookupName,Contact_City,Contact_Email,Contact_Mobile_Phone,Contact_Postal_Code,Contact_Province_State,Contact_Street FROM TOA.Work_Order WHERE ID =" + WorkId + "";
                globalContext.LogMessage("BuscarWorkOrder: \n " + queryString);
                clientRN.QueryCSV(clientInfoHeader, aPIAccessRequest, queryString, 1, "|", false, false, out CSVTableSet queryCSV, out byte[] FileData);
                foreach (CSVTable table in queryCSV.CSVTables)
                {
                    String[] rowData = table.Rows;
                    foreach (String data in rowData)
                    {
                        Char delimiter = '|';
                        String[] substrings = data.Split(delimiter);
                        Observaciones  = substrings[0];
                        Empresa = substrings[1];
                        Contacto = substrings[2];
                        city = substrings[3];
                        Correo = substrings[4];
                        PhoneCel = substrings[5];
                        postalCode = substrings[6];
                        streetAddress = substrings[8];
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ObtenerInfoWorkOrder" + ex.Message + "Detalle: " + ex.StackTrace);
            }
        }

        public bool Init()
        {
            try
            {
                bool result = false;
                EndpointAddress endPointAddr = new EndpointAddress(globalContext.GetInterfaceServiceUrl(ConnectServiceType.Soap));
                // Minimum required
                BasicHttpBinding binding = new BasicHttpBinding(BasicHttpSecurityMode.TransportWithMessageCredential);
                binding.Security.Message.ClientCredentialType = BasicHttpMessageCredentialType.UserName;
                binding.ReceiveTimeout = new TimeSpan(0, 10, 0);
                binding.MaxReceivedMessageSize = 1048576; //1MB
                binding.SendTimeout = new TimeSpan(0, 10, 0);
                // Create client proxy class
                clientRN = new RightNowSyncPortClient(binding, endPointAddr);
                // Ask the client to not send the timestamp
                BindingElementCollection elements = clientRN.Endpoint.Binding.CreateBindingElements();
                elements.Find<SecurityBindingElement>().IncludeTimestamp = false;
                clientRN.Endpoint.Binding = new CustomBinding(elements);
                // Ask the Add-In framework the handle the session logic
                globalContext.PrepareConnectSession(clientRN.ChannelFactory);
                if (clientRN != null)
                {
                    result = true;
                }

                return result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        public string RuleConditionInvoked(string ConditionName)
        {
            return string.Empty;
        }


    }




    [AddIn("Enviar Actividad", Version = "1.0.0.0")]
    public class WorkspaceAddInFactory : IWorkspaceComponentFactory2
    {
        IGlobalContext globalContext { get; set; }

        public IWorkspaceComponent2 CreateControl(bool inDesignMode, IRecordContext RecordContext)
        {
            return new WorkspaceAddIn(inDesignMode, RecordContext, globalContext);
        }


        public Image Image16
        {
            get { return Properties.Resources.AddIn16; }
        }


        public string Text
        {
            get { return "Enviar Actividad"; }
        }


        public string Tooltip
        {
            get { return "Enviar Actividad"; }
        }


        public bool Initialize(IGlobalContext GlobalContext)
        {
            this.globalContext = GlobalContext;
            return true;
        }


    }
}