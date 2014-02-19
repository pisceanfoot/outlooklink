using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OutlookLink
{
    public class OutlookTool
    {
        public static void OpenOutlookMail(string outlookLink)
        {
            outlookLink = outlookLink.Substring(8);

            try
            {

                Microsoft.Office.Interop.Outlook.Application app;
                Microsoft.Office.Interop.Outlook._NameSpace ns;

                app = new Microsoft.Office.Interop.Outlook.Application();
                ns = app.GetNamespace("MAPI");
                ns.Logon(null, null, false, false);

                Microsoft.Office.Interop.Outlook.MailItem mailItem = (Microsoft.Office.Interop.Outlook.MailItem)
                    ns.GetItemFromID(outlookLink);
                mailItem.Display();
            }
            catch (Exception ex)
            {

            }
        }
    }
}
