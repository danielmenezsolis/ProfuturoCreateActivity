using System;
using System.AddIn;
using System.Collections.Generic;
using System.Drawing;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Windows.Forms;
using Newtonsoft.Json;
using ProfuturoCreateActivity.POCRNService;
using RestSharp;
using RestSharp.Authenticators;
using RightNow.AddIns.AddInViews;
using RightNow.AddIns.Common;

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
        private string Regimen { get; set; }
        private string CURP { get; set; }
        private string Empresa { get; set; }
        private string Contacto { get; set; }
        private string PhoneCel { get; set; }
        private string PhoneHome { get; set; }
        private string Correo { get; set; }
        private string WorkOrderId { get; set; }
        private string Afore { get; set; }
        private string X { get; set; }
        private string Y { get; set; }
        private string TimeSlot { get; set; }
        private string WODate { get; set; }




        private RightNowSyncPortClient clientRN { get; set; }
        private IRecordContext recordContext { get; set; }
        private IGlobalContext globalContext { get; set; }
        private IIncident Incident { get; set; }
        private int IncidentID { get; set; }
        public IGenericObject SWorkOrder;
        public ICustomObject _workOrderRecord { get; set; }

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
                SWorkOrder = recordContext.GetWorkspaceRecord("TOA$Work_Order") as IGenericObject;
                if (Init())
                {


                    if (SWorkOrder != null)
                    {
                        recordContext.ExecuteEditorCommand(EditorCommand.Save);
                        WorkOrderId = SWorkOrder.Id.ToString();
                        ObtenerInfoWorkOrder(WorkOrderId);
                        CrearActividad();
                        recordContext.ExecuteEditorCommand(EditorCommand.Save);
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
            if (ActionName == "PopularContacto")
            {
                recordContext.ExecuteEditorCommand(EditorCommand.Save);
                PopulateContactDetails();
            }
            if (ActionName == "ValidaFecha")
            {
                _workOrderRecord = recordContext.GetWorkspaceRecord(recordContext.WorkspaceTypeName) as ICustomObject;
                IList<IGenericField> fields = _workOrderRecord.GenericFields;
                foreach (IGenericField field in fields)
                {
                    if (field.Name == "WO_Date")
                    {
                        DateTime WDate = Convert.ToDateTime(field.DataValue.Value);
                        if (WDate.Date < DateTime.Now.Date)
                        {
                            MessageBox.Show("La fecha de la cita no puede ser menor al día de hoy");
                            field.DataValue.Value = DateTime.Now.Date;
                        }
                    }
                }
            }

        }

        public void PopulateContactDetails()
        {
            try
            {
                IContact contactRecord = recordContext.GetWorkspaceRecord(WorkspaceRecordType.Contact) as IContact;
                _workOrderRecord = recordContext.GetWorkspaceRecord(recordContext.WorkspaceTypeName) as ICustomObject;
                IList<IGenericField> fields = _workOrderRecord.GenericFields;
                foreach (IGenericField field in fields)
                {
                    switch (field.Name)
                    {

                        case "Contact_Street":
                            if (contactRecord != null && contactRecord.AddrStreet != null)
                            {
                                field.DataValue.Value = contactRecord.AddrStreet;
                            }
                            break;
                        case "Contact_Province_State":
                            if (contactRecord != null && contactRecord.AddrProvID != null)
                            {
                                field.DataValue.Value = contactRecord.AddrProvID;
                            }
                            break;
                        case "Contact_City":
                            if (contactRecord != null && contactRecord.AddrCity != null)
                            {
                                field.DataValue.Value = contactRecord.AddrCity;
                            }
                            break;
                        case "Contact_Postal_Code":
                            if (contactRecord != null && contactRecord.AddrPostalCode != null)
                            {
                                field.DataValue.Value = contactRecord.AddrPostalCode;
                            }
                            break;
                        case "Contact_Email":
                            if (contactRecord != null && contactRecord.EmailAddr != null)
                            {
                                field.DataValue.Value = contactRecord.EmailAddr;
                            }
                            break;
                        case "Contact_Phone":
                            if (contactRecord != null && contactRecord.PhHome != null)
                            {
                                field.DataValue.Value = contactRecord.PhHome;
                            }
                            else if (contactRecord != null && contactRecord.PhOffice != null)
                            {
                                field.DataValue.Value = contactRecord.PhOffice;
                            }
                            break;
                        case "Contact_Mobile_Phone":
                            if (contactRecord != null && contactRecord.PhMobile != null)
                            {
                                field.DataValue.Value = contactRecord.PhMobile;
                            }
                            break;


                    }
                    recordContext.RefreshWorkspace();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Populate:" + ex.Message + " Det: " + ex.StackTrace);
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
                body += "\"apptNumber\":\"" + WorkOrderId + "\",";
                body += "\"resourceId\":\"z13m\",";
                body += "\"date\":\"" + WODate + "\",";
                body += "\"timeSlot\":\"" + TimeSlot + "\",";
                body += "\"activityType\":\"03\",";
                body += "\"customerName\":\"" + Contacto + "\",";
                body += "\"Apellido Paterno\":\"" + ApellidoPaterno + "\",";
                body += "\"Apellido Materno\":\"" + ApellidoMaterno + "\",";
                body += "\"streetAddress\":\"" + streetAddress + "\",";
                body += "\"city\":\"" + city + "\",";
                body += "\"postalCode\":\"" + postalCode + "\",";
                body += "\"stateProvince\":\"" + stateProvince + "\",";
                body += "\"NSS\":\"" + NSS + "\",";
                body += "\"Observaciones\":\"" + Observaciones + "\",";
                body += "\"Meses en su afore actual\":\"" + Meses + "\",";
                body += "\"Tipo de cartera\":\"Cedidos\",";
                body += "\"Afore\":\"" + Afore + "\",";
                body += "\"Régimen\":\"" + Regimen + "\",";
                body += "\"Curp\":\"" + CURP + "\",";
                body += "\"Nombre empresa\":\"" + Empresa + "\",";
                body += "\"customerPhone\":\"" + PhoneHome + "\",";
                body += "\"customerCell\":\"" + PhoneCel + "\",";
                if (!string.IsNullOrEmpty(X) && !string.IsNullOrEmpty(Y))
                {
                    body += "\"latitude\":" + X + ",";
                    body += "\"longitude\":" + Y + ",";
                }
                body += "\"customerEmail\":\"" + Correo + "\"";
                body += "}";
                globalContext.LogMessage("Crear Actividad: \n " + body);
                request.AddParameter("application/json", body, ParameterType.RequestBody);

                IRestResponse response = client.Execute(request);
                var content = response.Content; // raw content as string
                if (content.Contains("Bad"))
                {
                    MessageBox.Show("No se creó actividad, Motivo: " + content);
                }
                else if (content.Contains("activityId"))
                {
                    RootObject rootObject = JsonConvert.DeserializeObject<RootObject>(response.Content);
                    if (rootObject.activityId > 0)
                    {
                        MessageBox.Show("Se creó cita con número: " + rootObject.activityId.ToString());
                        _workOrderRecord = recordContext.GetWorkspaceRecord(recordContext.WorkspaceTypeName) as ICustomObject;

                        IList<IGenericField> fields = _workOrderRecord.GenericFields;

                        foreach (IGenericField field in fields)
                        {
                            if (field.Name == "External_ID")
                            {
                                field.DataValue.Value = rootObject.activityId.ToString();
                            }

                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("UpdatePaxPrice:" + ex.Message + " Det: " + ex.StackTrace);
            }

        }
        private void ObtenerInfoWorkOrder(string WorkId)
        {
            try
            {

                ClientInfoHeader clientInfoHeader = new ClientInfoHeader();
                APIAccessRequestHeader aPIAccessRequest = new APIAccessRequestHeader();
                clientInfoHeader.AppID = "Query Example";
                String queryString = "SELECT case_note,Contact.Name.First,  Contact_City,Contact_Email,Contact_Mobile_Phone, Contact_Postal_Code,Contact_State.name StateName  ,Contact_Street,Time_Slot.name TimeSlot, contact.CustomFields.c.org_name,  contact.CustomFields.c.curp,contact.CustomFields.c.nss, contact.CustomFields.c.afore.LookupName AS afore,  contact.CustomFields.c.regimen.LookupName AS regimen,X,Y,contact.CustomFields.c.amaterno,contact.CustomFields.c.fechacedido,contact_phone,Contact.Name.Last,WO_Date  FROM TOA.Work_Order  WHERE ID =" + Convert.ToInt32(WorkId) + "";
                globalContext.LogMessage("BuscarWorkOrder: \n " + queryString);
                clientRN.QueryCSV(clientInfoHeader, aPIAccessRequest, queryString, 1, "|", false, false, out CSVTableSet queryCSV, out byte[] FileData);
                foreach (CSVTable table in queryCSV.CSVTables)
                {
                    String[] rowData = table.Rows;
                    foreach (String data in rowData)
                    {
                        Char delimiter = '|';
                        String[] substrings = data.Split(delimiter);
                        Observaciones = substrings[0];
                        Contacto = substrings[1];
                        city = substrings[2];
                        Correo = substrings[3];
                        PhoneCel = substrings[4];
                        postalCode = substrings[5];
                        stateProvince = substrings[6];
                        streetAddress = substrings[7];
                        TimeSlot = substrings[8];
                        Empresa = substrings[9];
                        CURP = substrings[10];
                        NSS = substrings[11];
                        Afore = substrings[12];
                        Regimen = substrings[13];
                        X = substrings[14];
                        Y = substrings[15];
                        ApellidoMaterno = substrings[16];

                        Meses = GetMeses(string.IsNullOrEmpty(substrings[17]) ? DateTime.Now : Convert.ToDateTime(substrings[17]));
                        PhoneHome = substrings[18];
                        ApellidoPaterno = substrings[19];
                        WODate = substrings[20];

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ObtenerInfoWorkOrder" + ex.Message + "Detalle: " + ex.StackTrace);
            }
        }
        public string GetMeses(DateTime FechaCedido)
        {
            double totalMonths = 0;
            totalMonths = Math.Abs((FechaCedido.Month - DateTime.Now.Month) + 12 * (FechaCedido.Year - DateTime.Now.Year));
            //totalMonths = Math.Floor(totalMonths);
            return totalMonths.ToString();
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
public class Link
{
    public string rel { get; set; }
    public string href { get; set; }
}

public class RequiredInventories
{
    public List<Link> links { get; set; }
}

public class Link2
{
    public string rel { get; set; }
    public string href { get; set; }
}

public class LinkedActivities
{
    public List<Link2> links { get; set; }
}

public class Link3
{
    public string rel { get; set; }
    public string href { get; set; }
}

public class ResourcePreferences
{
    public List<Link3> links { get; set; }
}

public class Link4
{
    public string rel { get; set; }
    public string href { get; set; }
}

public class WorkSkills
{
    public List<Link4> links { get; set; }
}

public class Link5
{
    public string rel { get; set; }
    public string href { get; set; }
}

public class RootObject
{
    public int activityId { get; set; }
    public string resourceId { get; set; }
    public int resourceInternalId { get; set; }
    public string date { get; set; }
    public string apptNumber { get; set; }
    public string recordType { get; set; }
    public string status { get; set; }
    public string activityType { get; set; }
    public string workZone { get; set; }
    public int duration { get; set; }
    public int travelTime { get; set; }
    public string timeSlot { get; set; }
    public string customerPhone { get; set; }
    public string customerEmail { get; set; }
    public string customerCell { get; set; }
    public string streetAddress { get; set; }
    public string city { get; set; }
    public string postalCode { get; set; }
    public string stateProvince { get; set; }
    public string language { get; set; }
    public string languageISO { get; set; }
    public string timeZone { get; set; }
    public string timeZoneIANA { get; set; }
    public string serviceWindowStart { get; set; }
    public string serviceWindowEnd { get; set; }
    public string startTime { get; set; }
    public string endTime { get; set; }
    public string timeOfBooking { get; set; }
    public string resourceTimeZone { get; set; }
    public string resourceTimeZoneIANA { get; set; }
    public int resourceTimeZoneDiff { get; set; }
    public string NSS { get; set; }
    public string __invalid_name__Apellido_Paterno { get; set; }
    public string __invalid_name__Apellido_Materno { get; set; }
    public string Curp { get; set; }
    public string Observaciones { get; set; }
    public string __invalid_name__Meses_en_su_afore_actual { get; set; }
    public string __invalid_name__Nombre_empresa { get; set; }
    public string __invalid_name__Tipo_de_cartera { get; set; }
    public string Afore { get; set; }
    public string __invalid_name__Régimen { get; set; }
    public RequiredInventories requiredInventories { get; set; }
    public LinkedActivities linkedActivities { get; set; }
    public ResourcePreferences resourcePreferences { get; set; }
    public WorkSkills workSkills { get; set; }
    public List<Link5> links { get; set; }
}