
using System.Diagnostics;
using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Administration;

namespace SynContactoADPortal.Features.SynContactoADPortal
{
    /// <summary>
    /// This class handles events raised during feature activation, deactivation, installation, uninstallation, and upgrade.
    /// </summary>
    /// <remarks>
    /// The GUID attached to this class may be used during packaging and should not be modified.
    /// </remarks>

    [Guid("39936280-8aca-4a28-af54-1ff209283478")]
    public class SynContactoADPortalEventReceiver : SPFeatureReceiver
    {
        // Uncomment the method below to handle the event raised after a feature has been activated.

        private const string JobName = JOBSynContactoADPortal.JobTitle;
            //"AdP v6.0: Job de sincronizacion Active Directory -> SharePoint List";
        
        private void RemoveTimerJob()
        {
            foreach (SPJobDefinition job in SPFarm.Local.TimerService.JobDefinitions)
            {
                if (job.Name == JobName)
                {
                    job.Delete();
                }
            }
        }

        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                // remove timer job
                RemoveTimerJob();

                foreach (SPWebApplication webApplication in SPWebService.AdministrationService.WebApplications)
                {
                    if (webApplication.Status == SPObjectStatus.Online && webApplication.IsAdministrationWebApplication)
                    {
                        // create timer job
                        try
                        {
                            JOBSynContactoADPortal job = new JOBSynContactoADPortal(JobName, webApplication, null, SPJobLockType.None);
                            SPDailySchedule schedule = new SPDailySchedule();
                            schedule.BeginHour = 2;
                            schedule.BeginMinute = 0;
                            schedule.EndHour = 3;
                            job.Schedule = schedule;
                            job.Update();
                            break;
                        }
                        catch (Exception ex)
                        {
                            // error creating job
                            using (EventLog eventLog = new EventLog("Application"))
                            {
                                eventLog.Source = "Application";
                                eventLog.WriteEntry("Error en SyncContactoADP: " + ex.Message.ToString() + " \n " + ex.Data.ToString(), EventLogEntryType.Error, 6666, 1);
                            }
                        }

                    }
                }
            });
        }

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                // remove timer job
                RemoveTimerJob();
            });
        }
    }
}
