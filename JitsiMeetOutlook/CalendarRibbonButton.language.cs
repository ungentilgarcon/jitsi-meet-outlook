using System.Text.Json;

namespace JitsiMeetOutlook
{
    public partial class CalendarRibbonButton
    {
        private void setLanguage()
        {
            this.JitsiMeet.Label = "VDV Visio";
            this.buttonNewJitsiMeeting.Label = Globals.ThisAddIn.getElementTranslation("calendarRibbonButton", "buttonNewJitsiMeeting");
        }
    }
}
