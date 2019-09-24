using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint.Utilities;
using System.DirectoryServices.ActiveDirectory;
using System.Security.Principal;
using System.Data.SqlClient;
using System.Data;
using System.Diagnostics;


namespace SynContactoADPortal
{

    //JOB que actualiza los contactos del portal con los del directorio activo y en caso no se encuentre en el portal crea el contacto en el portal!
    //Autores: Felix Vargas & Gabriel Villayzan
    //SynContactoADPortal v.11.0 - 05.03.18 - Joshua Suasnabar
    //Se cambio la conexión de planilla para consultar la lista CTE del sitio HR de SAP
    //Se reconstruyo todo el codigo para que limpie la lista y vuelva a escribir solo los usuarios activos (tanto en SAP y con datos en el AD)

    class JOBSynContactoADPortal : SPJobDefinition
    {

        #region Locales
        
        public const string JobTitle = "AdP v11.2: Job de sincronizacion Active Directory -> SharePoint List";
        private const string JobDescription = "AdP v11.2: Timer Job que sinbcroniza contactos del Directorio Activo(anexo, fijo, cel y rpm), hacia el Portal --- Requerimientos Básicos: La lista debe llamarse 'Directorio Telefónico AdP' y trabaja minimamente con los nombres internos de la lista tipo contacto (sincroniza con OUTLOOK, despues de migracion GMD y con DNI en 'buscapersonas')";
        public static string rutaRaiz = "https://portaladp";
        public static string rutaHR = "https://portaladp/hr";

        #endregion

        #region Constructor

        public override string Description
        {
            get { return JobDescription; }
        }
        
        public JOBSynContactoADPortal() : base()
        {
            this.Title = JobTitle;
        }
        public JOBSynContactoADPortal(string jobname, SPWebApplication webApplication, SPServer server, SPJobLockType lockType) : base(jobname, webApplication, server, lockType)
        {
            this.Title = JobTitle;
        }
        public JOBSynContactoADPortal(string jobName, SPService service, SPServer server, SPJobLockType targetType) : base(jobName, service, server, targetType)
        {
        }
        public JOBSynContactoADPortal(string jobName, SPWebApplication webApplication)
            : base(jobName, webApplication, null, SPJobLockType.ContentDatabase)
        {
            this.Title = JobTitle;
        }
        
        #endregion

        #region Clases
        public class Usuario
        {
            public string displayName { get; set; }
            public string company { get; set; }
            public string l { get; set; }
            public string title { get; set; }
            public string homePhone { get; set; }
            public string mobile { get; set; }
            public string facsimileTelephoneNumber { get; set; }
            public string telephoneNumber { get; set; }
            public string email { get; set; }
            public string pager { get; set; }
            public string postalCode { get; set; }
            public string physicalDeliveryOfficeName { get; set; }
            public string aeropuerto { get; set; }

        }
        #endregion

        #region Metodos

        //Metodo reutilizable que permite consultar diferentes lista solo cambios el parametro nombreLista
        public static SPListItemCollection ConsultarLista(string nombreLista, string ruta)
        {
            String urlListas = ruta;
            SPSite nuevaColeccion = new SPSite(urlListas);
            SPWeb contenidoListas = nuevaColeccion.OpenWeb();
            SPList miLista = contenidoListas.Lists[nombreLista];
            SPListItemCollection ListaItems = miLista.Items;

            return ListaItems;
        }

        public SearchResultCollection conectarAD()
        {
            //consultamos el Active Directory 
            DirectoryContext dc = new DirectoryContext(DirectoryContextType.Domain, Environment.UserDomainName);
            Domain domain = Domain.GetDomain(dc);
            DirectoryEntry de = domain.GetDirectoryEntry();
            DirectorySearcher deSearch = new DirectorySearcher(de);
            deSearch.Filter = "(&(objectClass=user)(objectCategory=person)(mail=*)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))";
            deSearch.PropertiesToLoad.Add("mail");
            SearchResultCollection results = deSearch.FindAll();

            return results;
        }


        public List<Usuario> obtenerUsuariosAD()
        {
            //variables
            List<Usuario> listaUsuariosAD = new List<Usuario>();

            SearchResultCollection results = conectarAD();

            // Obtenemos todos los usuarios del AD
            foreach (SearchResult srUser in results)
            {
                DirectoryEntry deUser = srUser.GetDirectoryEntry();
                using (new SPMonitoredScope("Get displayName and company and l and title and facsimileTelephoneNumber and telephoneNumber and homePhone and mobile and objectClass From AD"))
                {
                    Usuario user = new Usuario();
                    //obtenemos el displayName
                    user.displayName = (deUser.InvokeGet("displayName") == null)? "" : deUser.InvokeGet("displayName").ToString();
                    //obtenemos la ciudad (City)
                    user.l = (deUser.InvokeGet("L") == null) ? "" : deUser.InvokeGet("l").ToString();
                    //obtenemos el puesto 
                    user.title = (deUser.InvokeGet("title") == null) ? "" : deUser.InvokeGet("title").ToString();
                    //obtenemos el Fax
                    user.facsimileTelephoneNumber = (deUser.InvokeGet("facsimileTelephoneNumber") == null) ? "" : deUser.InvokeGet("facsimileTelephoneNumber").ToString();
                    //obtenemos el anexo
                    user.telephoneNumber = (deUser.InvokeGet("telephoneNumber") == null) ? "" : deUser.InvokeGet("telephoneNumber").ToString();
                    //obtenemos el celular
                    user.homePhone = (deUser.InvokeGet("homePhone") == null) ? "" : deUser.InvokeGet("homePhone").ToString();
                    //obtenemos el RPM
                    user.mobile = (deUser.InvokeGet("mobile") == null) ? "" : deUser.InvokeGet("mobile").ToString();
                    //obtenemos el mail
                    user.email = (deUser.InvokeGet("mail") == null) ? "" : deUser.InvokeGet("mail").ToString();
                    //obtenemos el DNI
                    user.pager = (deUser.InvokeGet("pager") == null) ? "" : deUser.InvokeGet("pager").ToString();

                    listaUsuariosAD.Add(user);
                }
            }
            return listaUsuariosAD;
        }

        public void actualizarADconSAP(SPListItemCollection listaPlanilla)
        {

            //variables
            SearchResultCollection listaAD = conectarAD();

            try
            {
                foreach (SearchResult srUser in listaAD)
                {
                    DirectoryEntry deUser = srUser.GetDirectoryEntry();

                    if (deUser.InvokeGet("pager") != null)
                    {
                            foreach (SPListItem user in listaPlanilla)
                            {
                                string dniPlanilla = user["ows_DocumentoNumero"].ToString();
                                string dniAD = deUser.InvokeGet("pager").ToString();

                                if (dniAD.Equals(dniPlanilla))
                                {
                                    string puesto = "";
                                    string aeropuerto = "";
                                    string sede = "";
                                    string area = "";

                                    puesto= (user["ows_Puesto"] == null) ? "" : user["ows_Puesto"].ToString();
                                    sede = (user["ows_Sede"] == null) ? "" : user["ows_Sede"].ToString();
                                    area = (user["ows_Area"] == null) ? "" : user["ows_Area"].ToString();
                                    switch (sede)
                                    {
                                        case "AP01": aeropuerto = "LIMA"; break;
                                        case "AP02": aeropuerto = "AEROPUERTO TRUJILLO"; break;
                                        case "AP03": aeropuerto = "AEROPUERTO TARAPOTO"; break;
                                        case "AP04": aeropuerto = "AEROPUERTO PUCALLPA"; break;
                                        case "AP05": aeropuerto = "AEROPUERTO TALARA"; break;
                                        case "AP06": aeropuerto = "AEROPUERTO TUMBES"; break;
                                        case "AP07": aeropuerto = "AEROPUERTO CAJAMARCA"; break;
                                        case "AP08": aeropuerto = "AEROPUERTO CHACHAPOYAS"; break;
                                        case "AP09": aeropuerto = "AEROPUERTO ANTA"; break;
                                        case "AP10": aeropuerto = "AEROPUERTO PIURA"; break;
                                        case "AP11": aeropuerto = "AEROPUERTO PISCO"; break;
                                        case "AP12": aeropuerto = "AEROPUERTO CHICLAYO"; break;
                                        case "AP13": aeropuerto = "AEROPUERTO IQUITOS"; break;
                                        default    : aeropuerto = "ADP"; break;
                                    }
                                    if (puesto != "")
                                    {
                                        deUser.Properties["description"].Value = puesto;
                                        deUser.Properties["title"].Value = puesto;
                                    }
                                    if (aeropuerto != "")
                                    {
                                        deUser.Properties["physicalDeliveryOfficeName"].Value = aeropuerto;
                                        deUser.Properties["postalCode"].Value = aeropuerto;
                                    }
                                    if (area != "")
                                    {
                                        deUser.Properties["department"].Value = area;
                                    }
                                    
                                    
                                    deUser.CommitChanges();
                                }
                            }
                         
                        
                    }
                }
            }
            catch (Exception ex)
            {
                using (EventLog eventLog = new EventLog("Application"))
                {
                    eventLog.Source = "Application";
                    eventLog.WriteEntry("Error en actualización: " + ex.Message.ToString() + " \nDescripcion: " + ex.ToString()+ "\n" + ex.Data.ToString(), EventLogEntryType.Error, 6666, 1);
                }
            }
        }


        public bool verificarExiste(string correo, List<Usuario> listaPortal)
        {
            bool existe;
            return existe = listaPortal.Exists(element => element.email.ToLower() == correo.ToLower());
        }

        public void limpiarListaDirectorio(string site, string list)
        {
            using (SPSite spSite = new SPSite(site))
            {
                using (SPWeb spWeb = spSite.OpenWeb())
                {
                    StringBuilder deletebuilder = BatchCommand(spWeb.Lists[list]);
                    spSite.RootWeb.ProcessBatchData(deletebuilder.ToString());
                }
            }
        }

        private static StringBuilder BatchCommand(SPList spList)
        {
            StringBuilder deletebuilder = new StringBuilder();
            deletebuilder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Batch>");
            string command = "<Method><SetList Scope=\"Request\">" + spList.ID +
                "</SetList><SetVar Name=\"ID\">{0}</SetVar><SetVar Name=\"owsfileref\">{1}</SetVar><SetVar Name=\"Cmd\">Delete</SetVar></Method>";

            foreach (SPListItem item in spList.Items)
            {
                deletebuilder.Append(string.Format(command, item.ID.ToString(), item["FileRef"].ToString()));
            }
            deletebuilder.Append("</Batch>");
            return deletebuilder;
        }


        public void crearUsuariosEnDirectorioDesdeAD(SPListItemCollection listaDirectorio, List<Usuario> listaAD)
        {
            foreach (Usuario user in listaAD)
            {
                if (user.email.ToString() != "" && (user.homePhone.ToString() != "" || user.mobile.ToString() != "" || user.pager.ToString() != "" || user.facsimileTelephoneNumber.ToString() != ""))
                {
                    SPListItem item = listaDirectorio.Add();

                    item["ows_Title"] = user.displayName.ToString();
                    if (user.facsimileTelephoneNumber != "" && user.telephoneNumber != "")
                    {
                        item["ows_WorkPhone"] = user.facsimileTelephoneNumber.ToString() + " / " + user.telephoneNumber.ToString(); // Fijo
                    }
                    else if (user.facsimileTelephoneNumber != "" && user.telephoneNumber == "")
                    {
                        item["ows_WorkPhone"] = user.facsimileTelephoneNumber.ToString(); // Fijo
                    }
                    else if (user.facsimileTelephoneNumber == "" && user.telephoneNumber != "")
                    {
                        item["ows_WorkPhone"] = user.telephoneNumber.ToString(); // Fijo
                    }
                    item["ows_CellPhone"] = user.homePhone.ToString();
                    item["ows_HomePhone"] = user.mobile.ToString();
                    item["ows_Email"] = user.email.ToString();
                    item["ows_PagerNumber"] = user.pager.ToString();
                    item.Update();
                    }
            }
        }

        public void actualizarDirectorioConPlanilla(SPListItemCollection listaDirectorio, SPListItemCollection listaPlanilla)
        {
            foreach (SPListItem user in listaPlanilla)
            {
                foreach (SPListItem item in listaDirectorio)
                {
                    if (item["ows_Email"] != null && item["ows_Email"].ToString() != "" && item["ows_PagerNumber"] != null)
                    {
                            if (string.Compare(user["ows_DocumentoNumero"].ToString().Trim(), item["ows_PagerNumber"].ToString().Trim(), true) == 0)
                            {
                                //actualizamos el cargo
                                item["ows_JobTitle"] = user["ows_Puesto"].ToString();
                                item["ows_Gerencia"] = user["ows_Puesto"].ToString(); 
                                item["ows_FullName"] = string.Concat(user["ows_ApellidoPaterno"].ToString()+" "+
                                                                     user["ows_ApellidoMaterno"].ToString()+", "+
                                                                     user["ows_Nombres"].ToString()); 
                                switch(user["ows_Puesto"].ToString())
                                    {
                                        case "AP01": item["ows_WorkCity"] = "LIMA"; break;
                                        case "AP02": item["ows_WorkCity"] = "AEROPUERTO TRUJILLO"; break;
                                        case "AP03": item["ows_WorkCity"] = "AEROPUERTO Tarapoto"; break;
                                        case "AP04": item["ows_WorkCity"] = "AEROPUERTO PUCALLPA"; break;
                                        case "AP05": item["ows_WorkCity"] = "AEROPUERTO TALARA"; break;
                                        case "AP06": item["ows_WorkCity"] = "AEROPUERTO TUMBES"; break;
                                        case "AP07": item["ows_WorkCity"] = "AEROPUERTO CAJAMARCA"; break;
                                        case "AP08": item["ows_WorkCity"] = "AEROPUERTO CHACHAPOYAS"; break;
                                        case "AP09": item["ows_WorkCity"] = "AEROPUERTO ANTA"; break;
                                        case "AP10": item["ows_WorkCity"] = "AEROPUERTO PIURA"; break;
                                        case "AP11": item["ows_WorkCity"] = "AEROPUERTO PISCO"; break;
                                        case "AP12": item["ows_WorkCity"] = "AEROPUERTO CHICLAYO"; break;
                                        case "AP13": item["ows_WorkCity"] = "AEROPUERTO IQUITOS"; break;
                                    }    
                                    
                                item.Update();
                            }
                        }
                    }
                }
        }

        #endregion

        #region Execute
        public override void Execute(Guid targetInstanceId)
        {
            //consultamos la lista del portal que nos interesan y el AD
            SPListItemCollection listaPlanilla = ConsultarLista("CTE_Trabajadores", rutaHR);
            SPListItemCollection listaDirectorio = ConsultarLista("Directorio Telefónico AdP",rutaRaiz);
            

            actualizarADconSAP(listaPlanilla);
            List<Usuario> listaADactualizada = obtenerUsuariosAD();
            limpiarListaDirectorio(rutaRaiz,"Directorio Telefónico AdP");
            crearUsuariosEnDirectorioDesdeAD(listaDirectorio, listaADactualizada);
            actualizarDirectorioConPlanilla(listaDirectorio, listaPlanilla);
        }
        #endregion

    }
}
