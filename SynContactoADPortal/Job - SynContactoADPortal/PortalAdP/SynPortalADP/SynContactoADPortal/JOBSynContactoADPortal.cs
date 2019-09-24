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
    //SynContactoADPortal v.6.0

    class JOBSynContactoADPortal : SPJobDefinition
    {
        #region Locales
        
        
        public const string JobTitle = "AdP v8.0: Job de sincronizacion Active Directory -> SharePoint List";
        private const string JobDescription = "AdP v8.0: Timer Job que sinbcroniza contactos del Directorio Activo(anexo, fijo, cel y rpm), hacia el Portal --- Requerimientos Básicos: La lista debe llamarse 'Directorio Telefónico AdP' y trabaja minimamente con los nombres internos de la lista tipo contacto (sincroniza con OUTLOOK, despues de migracion GMD y con DNI en 'buscapersonas')";

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
            
            }
        #endregion

        #region Metodos

        public bool verificarExiste(string correo, List<Usuario> listaPortal)
        {
            bool existe;
            return existe = listaPortal.Exists(element => element.email.ToLower() == correo.ToLower());
        }
        public override void Execute(Guid targetInstanceId)
        {
            string webUrl = "https://portaladp/";

            using (SPWeb oWebsite = new SPSite(webUrl).OpenWeb())
            {

                oWebsite.AllowUnsafeUpdates = true;


                SPList lista = oWebsite.Lists["Directorio Telefónico AdP"];

                SPListItemCollection spcollLista = lista.Items;
                SPListItemCollection spcollLista1 = lista.Items;


                List<Usuario> listaUsuariosAD = new List<Usuario>();
                List<Usuario> listaUsuariosPortal = new List<Usuario>();
                DirectoryContext dc = new DirectoryContext(DirectoryContextType.Domain, Environment.UserDomainName);
                Domain domain = Domain.GetDomain(dc);
                DirectoryEntry de = domain.GetDirectoryEntry();

                DirectorySearcher deSearch = new DirectorySearcher(de);
                deSearch.Filter = "(&(objectClass=user)(objectCategory=person)(mail=*))";

                deSearch.PropertiesToLoad.Add("mail");
                SearchResult result = deSearch.FindOne();

                SearchResultCollection results = deSearch.FindAll();

                // Obtenemos todos los usuarios del AD
                foreach (SearchResult srUser in results)
                {
                    DirectoryEntry deUser = srUser.GetDirectoryEntry();

                    using (new SPMonitoredScope("Get displayName and company and l and title and facsimileTelephoneNumber and telephoneNumber and homePhone and mobile and objectClass From AD"))
                    {
                        Usuario user = new Usuario();
                        if (deUser.InvokeGet("displayName") == null)
                        {
                            user.displayName = "";
                        }
                        else
                        {
                            user.displayName = deUser.InvokeGet("displayName").ToString();
                        }

                        if (deUser.InvokeGet("company") == null)
                        {
                            user.company = "";
                        }
                        else
                        {
                            user.company = deUser.InvokeGet("company").ToString();
                        }

                        if (deUser.InvokeGet("L") == null)
                        {
                            user.l = "";
                        }
                        else
                        {
                            user.l = deUser.InvokeGet("l").ToString();
                        }

                        if (deUser.InvokeGet("title") == null)
                        {
                            user.title = "";
                        }
                        else
                        {
                            user.title = deUser.InvokeGet("title").ToString();
                        }
                        if (deUser.InvokeGet("facsimileTelephoneNumber") == null)
                        {
                            user.facsimileTelephoneNumber = "";
                        }
                        else
                        {
                            user.facsimileTelephoneNumber = deUser.InvokeGet("facsimileTelephoneNumber").ToString(); // anexo 513 nortel
                        }
                        if (deUser.InvokeGet("telephoneNumber") == null)
                        {
                            user.telephoneNumber = "";
                        }
                        else
                        {
                            user.telephoneNumber = deUser.InvokeGet("telephoneNumber").ToString(); // anexo 513 nortel
                        }
                        if (deUser.InvokeGet("homePhone") == null)
                        {
                            user.homePhone = "";
                        }
                        else
                        {
                            user.homePhone = deUser.InvokeGet("homePhone").ToString(); // cel completo
                        }
                        if (deUser.InvokeGet("mobile") == null)
                        {
                            user.mobile = "";
                        }
                        else
                        {
                            user.mobile = deUser.InvokeGet("mobile").ToString(); // rpm
                        }
                        if (deUser.InvokeGet("mail") == null)
                        {
                            user.mobile = "";
                        }
                        else
                        {
                            user.email = deUser.InvokeGet("mail").ToString();
                        }
                        // obtener DNI
                        if (deUser.InvokeGet("pager") == null)
                        {
                            user.pager = "";
                        }
                        else
                        {
                            user.pager = deUser.InvokeGet("pager").ToString();
                        }



                        listaUsuariosAD.Add(user);
                    }
                }

                // obtener usuarios Portal

                foreach (SPListItem item in spcollLista)
                {
                    Usuario user = new Usuario();
                    if (item["ows_Email"] == null || item["ows_Email"].ToString() == "")
                    {
                    }
                    else
                    {
                        if (item["ows_Title"] == null || item["ows_Title"].ToString() == "")
                        {
                            user.displayName = "";
                        }
                        else
                        {
                            user.displayName = item["ows_Title"].ToString();
                        }
                        //if (item["ows_Company"] == null || item["ows_Company"].ToString() == "")
                        //{
                        //    user.company = "";
                        //}
                        //else
                        //{
                        //    user.company = item["ows_Company"].ToString();
                        //}

                        if (item["ows_WorkPhone"] == null || item["ows_WorkPhone"].ToString() == "")
                        {
                            user.facsimileTelephoneNumber = "";
                        }
                        else
                        {
                            user.facsimileTelephoneNumber = item["ows_WorkPhone"].ToString();
                        }
                        if (item["ows_CellPhone"] == null || item["ows_CellPhone"].ToString() == "")
                        {
                            user.homePhone = "";
                        }
                        else
                        {
                            user.homePhone = item["ows_CellPhone"].ToString();
                        }
                        if (item["ows_HomePhone"] == null || item["ows_HomePhone"].ToString() == "")
                        {
                            user.mobile = "";
                        }
                        else
                        {
                            user.mobile = item["ows_HomePhone"].ToString();
                        }
                        if (item["ows_WorkCity"] == null || item["ows_WorkCity"].ToString() == "")
                        {
                            user.l = "";
                        }
                        else
                        {
                            user.l = item["ows_WorkCity"].ToString();
                        }
                        if (item["ows_JobTitle"] == null || item["ows_JobTitle"].ToString() == "")
                        {
                            user.title = "";
                        }
                        else
                        {
                            user.title = item["ows_JobTitle"].ToString();
                        }
                        if (item["ows_Email"] == null || item["ows_Email"].ToString() == "")
                        {
                            user.email = "";
                        }
                        else
                        {
                            user.email = item["ows_Email"].ToString();
                        }
                        listaUsuariosPortal.Add(user);
                        //DNI
                        if (item["ows_PagerNumber"] == null || item["ows_PagerNumber"].ToString() == "")
                        {
                            user.pager = "";
                        }
                        else
                        {
                            user.pager = item["ows_PagerNumber"].ToString();
                        }
                        listaUsuariosPortal.Add(user);
                    }

                }


                // Actualiza todos los Items de la Lista DIrectorio Telefonico!
                foreach (Usuario user in listaUsuariosAD)
                {

                    foreach (SPListItem item in spcollLista)
                    {

                        if (item["ows_Email"] == null || item["ows_Email"].ToString() == "")
                        {
                        }
                        else
                        {
                            if (string.Compare(user.email.ToString(), item["ows_Email"].ToString(), true) == 0)
                            {
                                if (item["ows_Title"] == null || item["ows_Title"].ToString() != user.displayName.ToString())
                                {
                                    item["ows_Title"] = user.displayName.ToString(); // Nombre
                                    item.Update();
                                }
                                //if (item["ows_Company"] == null || item["ows_Company"].ToString() != user.company.ToString())
                                //{
                                //    item["ows_Company"] = user.company.ToString(); // Empresa
                                //    item.Update();
                                //}

                                if (item["ows_WorkPhone"] == null || item["ows_WorkPhone"].ToString() != user.facsimileTelephoneNumber.ToString())
                                {
                                    if (user.facsimileTelephoneNumber != "" && user.telephoneNumber != "")
                                    {
                                        item["ows_WorkPhone"] = user.facsimileTelephoneNumber.ToString() + " / " + user.telephoneNumber.ToString(); // Fijo
                                        item.Update();
                                    }
                                    else if (user.facsimileTelephoneNumber != "" && user.telephoneNumber == "")
                                    {
                                        item["ows_WorkPhone"] = user.facsimileTelephoneNumber.ToString(); // Fijo
                                        item.Update();
                                    }
                                    else if (user.facsimileTelephoneNumber == "" && user.telephoneNumber != "")
                                    {
                                        item["ows_WorkPhone"] = user.telephoneNumber.ToString(); // Fijo
                                        item.Update();
                                    }
                                }

                                if (item["ows_CellPhone"] == null || item["ows_CellPhone"].ToString() != user.homePhone.ToString())
                                {
                                    item["ows_CellPhone"] = user.homePhone.ToString(); // movil
                                    item.Update();
                                }

                                if (item["ows_HomePhone"] == null || item["ows_HomePhone"].ToString() != user.mobile.ToString())
                                {
                                    item["ows_HomePhone"] = user.mobile.ToString(); // rpm
                                    item.Update();
                                }
                                //if (item["ows_WorkCity"] == null || item["ows_WorkCity"].ToString() != user.l.ToString())
                                //{
                                //    item["ows_WorkCity"] = user.l.ToString(); // ubicacion
                                //    item.Update();
                                //}
                                if (item["ows_JobTitle"] == null || item["ows_JobTitle"].ToString() != user.title.ToString())
                                {
                                    item["ows_JobTitle"] = user.title.ToString(); // cargo
                                    item.Update();
                                }
                                
                                if (item["ows_PagerNumber"] == null || item["ows_PagerNumber"].ToString() != user.pager.ToString())
                                {
                                    item["ows_PagerNumber"] = user.pager.ToString(); // DNI
                                    item.Update();
                                }
                                break;
                            }
                        }
                    }
                }

                // Capturamos los datos de la lista [Lista de Trabajadores] de RRHH
                String sSel = "select DNI,PUESTO_TRABAJO from V_TrabajadoresADP";
                SqlConnection sCnn =
                    new SqlConnection("data source = 192.168.57.19; initial catalog = OFIPLAN; user id = sharepoint; password = 5h4r3p01nt#");
            
                SqlDataAdapter da= new SqlDataAdapter(sSel, sCnn);
                DataTable dt = new DataTable();

                    da.Fill(dt);

                    List<Usuario> usuarios = new List<Usuario>();

                    foreach (DataRow dr in dt.Rows)
                    {
                        Usuario usuario = new Usuario();
                        string DNI = (string)dr["DNI"];
                        string Puesto = (string)dr["PUESTO_TRABAJO"];

                        usuario.pager = DNI;
                        usuario.title = Puesto;

                        usuarios.Add(usuario);

                    }


                // Actualiza todos los cargos de la Lista DIrectorio Telefonico!

                    try
                    {
                        foreach (Usuario user in usuarios)
                        {

                            foreach (SPListItem item in spcollLista)
                            {
                                if (item["ows_Email"] == null || item["ows_Email"].ToString() == "")
                                {
                                }
                                else
                                {
                                    if (item["ows_PagerNumber"] != null)
                                    {
                                        if (string.Compare(user.pager.ToString().Trim(), item["ows_PagerNumber"].ToString().Trim(), true) == 0)
                                        {
                                            if (item["ows_JobTitle"] == null || item["ows_JobTitle"].ToString() != user.title.ToString())
                                            {
                                                item["ows_JobTitle"] = user.title.ToString(); // cargo
                                                item.Update();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // error creating job
                        using (EventLog eventLog = new EventLog("Application"))
                        {
                            eventLog.Source = "Application";
                            eventLog.WriteEntry("Error en actualización: " + ex.Message.ToString() + " \n " + ex.Data.ToString(), EventLogEntryType.Error, 6666, 1);
                        }
                    }
                

                foreach (Usuario user in listaUsuariosAD)
                {
                    if (user.email.ToString() != "")
                    {
                        if ((verificarExiste(user.email.ToString(), listaUsuariosPortal) == false) && (user.telephoneNumber != "" || user.facsimileTelephoneNumber != "" || user.homePhone != "" || user.mobile != ""))
                        //if ((verificarExiste(user.email.ToString(), listaUsuariosPortal) == false) )
                        {
                            SPListItem item1 = spcollLista.Add();

                            item1["ows_Title"] = user.displayName.ToString();
                            //item1["ows_Company"] = user.company.ToString();
                            if (user.facsimileTelephoneNumber != "" && user.telephoneNumber != "")
                            {
                                item1["ows_WorkPhone"] = user.facsimileTelephoneNumber.ToString() + " / " + user.telephoneNumber.ToString(); // Fijo
                            }
                            else if (user.facsimileTelephoneNumber != "" && user.telephoneNumber == "")
                            {
                                item1["ows_WorkPhone"] = user.facsimileTelephoneNumber.ToString(); // Fijo

                            }
                            else if (user.facsimileTelephoneNumber == "" && user.telephoneNumber != "")
                            {
                                item1["ows_WorkPhone"] = user.telephoneNumber.ToString(); // Fijo

                            }

                            item1["ows_CellPhone"] = user.homePhone.ToString();
                            item1["ows_HomePhone"] = user.mobile.ToString();
                            //item1["ows_WorkCity"] = user.l.ToString();
                            //item1["ows_JobTitle"] = user.title.ToString();
                            item1["ows_Email"] = user.email.ToString();
                            item1["ows_PagerNumber"] = user.pager.ToString();
                            item1.Update();
                        }
                        else
                        {

                        }
                    }
                    else
                    {

                    }
                }
                oWebsite.AllowUnsafeUpdates = false;
            }
        }
        #endregion

    }
}
