using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace JitsiMeetOutlook
{
    public class Utils
    {
        public const string JitsiAppointmentPropertyName = "VisioVilleurbanneMeeting";

        private static string escapeDomain(string domain)
        {
            string escapedDomain = Regex.Escape(domain);
            if (!escapedDomain.EndsWith("/"))
            {
                escapedDomain += "/";
            }
            return escapedDomain;
        }

        public static string GetUrl(string oldBody, string domain)
        {
            return Regex.Match(oldBody, "http[s]*://" + escapeDomain(domain) + "[\\w\\/#%\\.=]+").Value;
        }

        public static bool SettingIsActive(string url, string setting)
        {
            return url.Contains("config." + setting + "=true");
        }


        public static string findRoomId(string body, string domain)
        {
            string roomId = Regex.Match(body, "(?<=" + escapeDomain(domain) + ")\\S+?(?=(#config|&config|\\s))").Value; // Match all non-blanks after jitsi url and before config or end
            return roomId;
        }

        public static bool IsJitsiAppointment(Outlook.AppointmentItem appointment, string domain)
        {
            if (appointment == null)
            {
                return false;
            }

            if (appointment.Location == "Jitsi Meet")
            {
                return true;
            }

            Outlook.UserProperty marker = appointment.UserProperties.Find(JitsiAppointmentPropertyName, false);
            if (marker != null && marker.Value is bool isJitsiMeeting && isJitsiMeeting)
            {
                return true;
            }

            return !string.IsNullOrEmpty(findRoomId(appointment.Body, domain));
        }

        public static void RunInThread(Action function)
        {
            try
            {
                ThreadStart s = new ThreadStart(() =>
                {
                    try
                    {
                        function();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("An Error occured within JitsiOutlook: " + ex.Message);
                    }
                });
                Thread ss = new Thread(s);
                ss.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An Error occured within JitsiOutlook: " + ex.Message);
            }
        }

        public static List<KeyValuePair<bool, string>> SplitToTextAndHyperlinks(string text)
        {
            var list = new List<KeyValuePair<bool, string>>();
            MatchCollection matches = Regex.Matches(text, "http[s]?:\\/\\/(?:[a-zA-Z]|[0-9]|[$-_@&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?<!\\.)");
            if (matches.Count == 0)
            {
                list.Add(new KeyValuePair<bool, string>(false, text));
            }
            var lastindex = 0;
            var index = 0;
            foreach (Match match in matches)
            {
                list.Add(new KeyValuePair<bool, string>(false, text.Substring(lastindex, match.Index - lastindex)));
                list.Add(new KeyValuePair<bool, string>(true, match.Value));
                lastindex = match.Index + match.Length;
                if (index == matches.Count - 1)
                {
                    list.Add(new KeyValuePair<bool, string>(false, text.Substring(lastindex, text.Length - lastindex)));
                }
                index++;
            }

            return list;
        }
    }
}
