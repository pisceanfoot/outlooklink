using System.Windows.Forms;
using Microsoft.Win32;

namespace OutlookLink
{
    public class RegisterTool
    {
        public static bool Register(string protocol, string application, string arguments, RegistryKey registry = null)
        {
            if (registry == null)
            {
                registry = Registry.CurrentUser;
            }

            RegistryKey cl = Registry.ClassesRoot.OpenSubKey(protocol);

            //if (cl != null && cl.GetValue("URL Protocol") != null && cl.GetValue("CustomUrlApplication") == null)
            //    if (System.Windows.Forms.MessageBox.Show("Protocol '" + protocol + "' is already registered. Do you wish to overwrite the current information?", "CustomURL", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.No)
            //        return false;

            try
            {
                RegistryKey r;
                r = registry.OpenSubKey("SOFTWARE\\Classes\\" + protocol, true);
                if (r == null)
                    r = registry.CreateSubKey("SOFTWARE\\Classes\\" + protocol);
                r.SetValue("", "URL: Protocol");
                r.SetValue("URL Protocol", "");
                r.SetValue("CustomUrlApplication", application);
                r.SetValue("CustomUrlArguments", arguments);

                r = registry.OpenSubKey("SOFTWARE\\Classes\\" + protocol + "\\DefaultIcon", true);
                if (r == null)
                    r = registry.CreateSubKey("SOFTWARE\\Classes\\" + protocol + "\\DefaultIcon");
                r.SetValue("", application);

                r = registry.OpenSubKey("SOFTWARE\\Classes\\" + protocol + "\\shell\\open\\command", true);
                if (r == null)
                    r = registry.CreateSubKey("SOFTWARE\\Classes\\" + protocol + "\\shell\\open\\command");

                r.SetValue("", application + " \"%1\"");


                // If 64-bit OS, also register in the 32-bit registry area. 
                if (registry.OpenSubKey("SOFTWARE\\Wow6432Node\\Classes") != null)
                {
                    r = registry.OpenSubKey("SOFTWARE\\Wow6432Node\\Classes\\" + protocol, true);
                    if (r == null)
                        r = registry.CreateSubKey("SOFTWARE\\Wow6432Node\\Classes\\" + protocol);
                    r.SetValue("", "URL: Protocol");
                    r.SetValue("URL Protocol", "");
                    r.SetValue("CustomUrlApplication", application);
                    r.SetValue("CustomUrlArguments", arguments);

                    r = registry.OpenSubKey("SOFTWARE\\Wow6432Node\\Classes\\" + protocol + "\\DefaultIcon", true);
                    if (r == null)
                        r = registry.CreateSubKey("SOFTWARE\\Wow6432Node\\Classes\\" + protocol + "\\DefaultIcon");
                    r.SetValue("", application);

                    r = registry.OpenSubKey("SOFTWARE\\Wow6432Node\\Classes\\" + protocol + "\\shell\\open\\command", true);
                    if (r == null)
                        r = registry.CreateSubKey("SOFTWARE\\Wow6432Node\\Classes\\" + protocol + "\\shell\\open\\command");

                    r.SetValue("", application + " \"%1\"");

                }

            }
            catch (System.UnauthorizedAccessException ex)
            {
                MessageBox.Show("You do not have permission to make changes to the registry!\n\nMake sure that you have administrative rights on this computer.", "CustomURL", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;

        }
    }
}
