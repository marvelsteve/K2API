using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;

namespace OGCIOCommonLibraryNet35
{
    public class NotificationUtil
    {
        private NotificationUtil()
        {
        }

        public static void SendEmail(int? documentid, int? processid, string from, string to, string subject, string emailTemplate, string hostname)
        {
            try
            {
                string text = ConstructEmailContent(emailTemplate, documentid, processid);
                if (string.IsNullOrEmpty(text)) return;

                MailMessage mm = new MailMessage(from, to, subject, text);
                SmtpClient smtp = new SmtpClient(hostname);
                smtp.Send(mm);
            }
            catch (Exception ex)
            {
                K2Util.LogError(ex.ToString());
            }
        }

        public static string ConstructEmailContent(string emailTemplate, int? documentid, int? processid)
        {
            try
            {
                string content = ConstructEmailContentForSmartObjects(emailTemplate, documentid, processid);
                content = ConstructEmailContentForProcessInformation(content, documentid, processid);
                return content;
            }
            catch (Exception ex)
            {
                K2Util.LogError(ex.ToString());
                return "error";
            }
        }

        private static string ConstructEmailContentForProcessInformation(string emailTemplate, int? documentid, int? processid)
        {
            if (processid == null) return emailTemplate;

            string text = emailTemplate;
            if (string.IsNullOrEmpty(text)) return "";

            string smoText = "@process.[";
            string documentIdText = "<documentid>";
            string processIdText = "<processid>";
            int start = 0;
            while (true)
            {
                int index = text.IndexOf(smoText, start, StringComparison.OrdinalIgnoreCase);
                if (index > -1)
                {
                    int index2 = text.IndexOf("]", index, StringComparison.OrdinalIgnoreCase);
                    if (index2 > -1)
                    {
                        string param = text.Substring(index, index2 + 1 - index);
                        if (documentid != null)
                        {
                            param = param.Replace(documentIdText, documentid.Value.ToString());
                        }
                        if (processid != null)
                        {
                            param = param.Replace(processIdText, processid.Value.ToString());
                        }
                        param = param.Replace("[", "");
                        param = param.Replace("]", "");

                        //object obj = K2Util.GetProcessInformationByPID();
                        object obj = null;
                        if (param.StartsWith("@process.", StringComparison.OrdinalIgnoreCase))
                        {
                            string[] arr = param.Split(".".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                            if (arr.Length >= 2)
                            {
                                int? pid = processid;

                                string type = arr[1];
                                string propName = "";
                                if (arr.Length >= 3) propName = arr[2];

                                obj = K2Util.GetProcessInformationByPID(processid.Value, type, propName);
                            }
                        }

                        text = text.Substring(0, index) + Convert.ToString(obj) + text.Substring(index2 + 1);

                        start = index;
                    }
                    else
                    {
                        break;
                    } // end if
                }
                else
                {
                    break;
                } // end if
            } // end while

            return text;
        }

        private static string ConstructEmailContentForSmartObjects(string emailTemplate, int? documentid, int? processid)
        {
            //string text = GetEmailTemplate();
            string text = emailTemplate;
            if (string.IsNullOrEmpty(text)) return "";

            string smoText = "@SMO.";
            string documentIdText = "<documentid>";
            string processIdText = "<processid>";
            int start = 0;
            while (true)
            {
                int index = text.IndexOf(smoText, start, StringComparison.OrdinalIgnoreCase);
                if (index > -1)
                {
                    int index2 = text.IndexOf("]", index, StringComparison.OrdinalIgnoreCase);
                    if (index2 > -1)
                    {
                        string param = text.Substring(index, index2 + 1 - index);
                        if (documentid != null)
                        {
                            param = param.Replace(documentIdText, documentid.Value.ToString());
                        }
                        if (processid != null)
                        {
                            param = param.Replace(processIdText, processid.Value.ToString());
                        }

                        object obj = K2Util.GetSmartObject(param, processid);

                        text = text.Substring(0, index) + Convert.ToString(obj) + text.Substring(index2 + 1);

                        start = index;
                    }
                    else
                    {
                        break;
                    } // end if
                }
                else
                {
                    break;
                } // end if
            } // end while

            return text;
        }

        /// <summary>
        /// Temp code for testing
        /// </summary>
        /// <returns></returns>
        public static string GetEmailTemplate()
        {
            string text = System.IO.File.ReadAllText(@"D:\Hugo\email_templates.txt");
            return text;
        }

    } // end class
}
