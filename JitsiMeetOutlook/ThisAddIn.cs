using JitsiMeetOutlook.Entities;
using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace JitsiMeetOutlook
{
    public partial class ThisAddIn
    {
        private readonly Lazy<JitsiApiService> _jitsiApiService = new Lazy<JitsiApiService>(() => new JitsiApiService());
        public JitsiApiService JitsiApiService => _jitsiApiService.Value;

        private MAPIFolder _calendarFolder;
        private Items _calendarFolderItems;
        private bool _calendarEventsHooked;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Near-zero work on UI thread - just schedule background initialization
            _ = Task.Run(InitializeAsync);
        }

        private async Task InitializeAsync()
        {
            try
            {
                await Task.Yield();
                checkFirstRunSettingsAsync();
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Startup background error: {ex}");
            }
        }

        private void EnsureCalendarEventsHooked()
        {
            if (_calendarEventsHooked) return;

            try
            {
                _calendarFolder = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                _calendarFolderItems = _calendarFolder.Items;
                _calendarFolderItems.ItemAdd += new ItemsEvents_ItemAddEventHandler(AppointmentAddedOrChanged);
                _calendarFolderItems.ItemChange += new ItemsEvents_ItemChangeEventHandler(AppointmentAddedOrChanged);
                _calendarEventsHooked = true;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"HookCalendarEvents error: {ex}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void checkFirstRunSettingsAsync()
        {
            // Quick check without file I/O first
            if (!IsFirstRunNeeded()) return;

            Task.Run(() =>
            {
                try
                {
                    if (isFirstRun() || isTxtOlderThanInstallation())
                    {
                        customSettingsOnFirstRun();
                    }
                }
                catch { /* ignore file errors */ }
            });
        }

        private bool IsFirstRunNeeded()
        {
            // Quick heuristic - avoid file I/O if marker exists and is recent
            try
            {
                string markerPath = firstRunMarkerFilePath();
                return !File.Exists(markerPath) ||
                       (DateTime.Now - File.GetLastWriteTime(markerPath)).TotalDays > 1;
            }
            catch { return true; }
        }

        private bool isFirstRun()
        {
            return !File.Exists(firstRunMarkerFilePath());
        }

        private bool isTxtOlderThanInstallation()
        {
            string mainDllPath = Path.Combine(Properties.Settings.Default.installDirectory, "Jitsi Meet Outlook Add-In.dll.config");

            FileInfo infoMainDll = new FileInfo(mainDllPath);
            FileInfo infoFirstRunTxt = new FileInfo(firstRunMarkerFilePath());

            return infoMainDll.LastWriteTime > infoFirstRunTxt.LastWriteTime;
        }

        private string firstRunMarkerFilePath()
        {
            string shadowCopyLocation = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string firstRunMarkerFile = Path.Combine(shadowCopyLocation, "firstrun.txt");

            return firstRunMarkerFile;
        }

        private void customSettingsOnFirstRun()
        {
            // Reset settings to those specified during install, and if none were specified, to default settings
            Properties.Settings.Default.Reset();

            // Create a file in the assembly shadow copy location to mark this code as executed
            string firstRunText = "This file is created the first time the add-in is launched.\n\nRemoval of this file will reset settings, whereafter it will be re-created on next launch of the add-in.";
            File.WriteAllText(firstRunMarkerFilePath(), firstRunText);
        }

        public void AppointmentAddedOrChanged(object appointment)
        {
            EnsureCalendarEventsHooked();

            // Make sure we only react to Appointments
            if (appointment is AppointmentItem)
            {
                AppointmentItem item = appointment as AppointmentItem;

                if (Utils.IsJitsiAppointment(item, Properties.Settings.Default.Domain))
                {
                    Utils.RunInThread(() =>
                    {
                        string recurrencePattern = null;
                        if (item.IsRecurring)
                        {
                            var recurrencebuilder = new Recurrence();
                            recurrencePattern = recurrencebuilder.buildRrule(item.GetRecurrencePattern());
                        }

                        var scheduledConference = new ConferenceSchedulerMessage
                        {

                            ConferenceName = Utils.findRoomId(item.Body, Properties.Settings.Default.Domain),
                            Start = item.StartUTC,
                            End = item.EndUTC,
                            Recurrance = recurrencePattern
                        };
                        JitsiApiService.ScheduleConference(scheduledConference);
                    });
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
