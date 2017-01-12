using System;
using System.Security.Principal;
using System.Runtime.InteropServices;
using System.Configuration;
using System.Text;

namespace VSOImpersonateWindowsUser_NS
{
    class VSOImpersonateWindowsUser
    {
        public class Impersonator : IDisposable
        {
            private WindowsImpersonationContext _impersonatedUser = null;
            private IntPtr _userHandle;
            public Impersonator()
            {
                _userHandle = new IntPtr(0);
                //Define user parameters for logon
                string userName = "liam harper";
                string userDomain = ConfigurationManager.AppSettings["MachineDomain"];
                string passwordB64 = "I0pQRjdBUlBCWVhQU1BWUVk=";
                byte[] passwordBytes = Convert.FromBase64String(passwordB64);
                string passwordFinal = Encoding.UTF8.GetString(passwordBytes);
                bool returnValue = LogonUser(userName, userDomain, passwordFinal, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, ref _userHandle);
                if (!returnValue)
                    throw new ApplicationException("Could not impersonate user.");

                WindowsIdentity windowsIndentity = new WindowsIdentity(_userHandle);
                _impersonatedUser = windowsIndentity.Impersonate();
            }
            #region IDisposable Members
            public void Dispose()
            {
                if (_impersonatedUser != null)
                {
                    _impersonatedUser.Undo();
                    CloseHandle(_userHandle);
                }
            }
            #endregion
            #region Interop imports/constants
            public const int LOGON32_LOGON_INTERACTIVE = 2;
            public const int LOGON32_LOGON_SERVICE = 3;
            public const int LOGON32_PROVIDER_DEFAULT = 0;
            [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
            public static extern bool LogonUser(String lpszUserName, String lpszDomain, String lpszPassword, int dwLogonType, int dwLogonProvider, ref IntPtr phToken);
            [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
            public extern static bool CloseHandle(IntPtr handle);
            #endregion
        }


    }
}
