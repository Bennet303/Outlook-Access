using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Reflection;

namespace Outlook_Access
{
    public abstract class AccessClass
    {
        protected Outlook.Application OutlookApplication { get; set; }
        protected Outlook.NameSpace OutlookNamespace { get; set; }
        protected Outlook.MAPIFolder OutlookFolder { get; set; }
        protected Outlook.Items OutlookFolderItems { get; set; }

        public void Connect(string pUsername, string pPassword, bool pShowDialog, bool pNewSession, string pNamespace, string pFolderID)
        {
            try
            {
                OutlookApplication = new Outlook.Application();
                OutlookNamespace = OutlookApplication.GetNamespace(pNamespace);
                object username = pUsername;
                object password = pPassword;

                if (username == null)
                {
                    username = Missing.Value;
                }
                if (password == null)
                {
                    password = Missing.Value;
                }

                OutlookNamespace.Logon(username, password, pShowDialog, pNewSession);

                if (pFolderID != null)
                {
                    OutlookFolder = OutlookNamespace.GetFolderFromID(pFolderID);
                }     
                // (else: OutlookFolder and) OutlookFolderItems has to be set in the inheriting class
            }
            catch (Exception e)
            {
                Console.WriteLine($"{e} Exception caught");
            }
        }
        public void Disconnect()
        {
            OutlookNamespace.Logoff();

            OutlookApplication = null;
            OutlookFolder = null;
            OutlookNamespace = null;
            OutlookFolderItems = null;
        } 
    }
}
