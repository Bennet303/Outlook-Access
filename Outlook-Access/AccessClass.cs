using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Reflection;
using System.Text;

namespace Outlook_Access
{
    public abstract class AccessClass
    {
        protected Outlook.Application OutlookApplication { get; set; }
        protected Outlook.NameSpace OutlookNamespace { get; set; }
        protected Outlook.MAPIFolder OutlookFolder { get; set; }
        protected Outlook.Items OutlookFolderItems { get; set; }

        protected void Connect(string pUsername, string pPassword, bool pShowDialog, bool pNewSession, string pNamespace, string pFolderID)
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

        //Builds a filter-string that can be used to search for items matching this filter
        protected static StringBuilder BuildFilterString(string pKeyword, string pValue, bool pIsSubstring, string pCriteria)
        {
            StringBuilder subjectFilter = new StringBuilder();
            if (pIsSubstring == true)
            {
                subjectFilter.Append("@SQL=");
                subjectFilter.Append(pCriteria);
                string valueCondition = String.Format(" like '%{0}%'", pValue);
                subjectFilter.Append(valueCondition);
            }
            else
            {
                string condition = String.Format("[{0}] = {1}", pKeyword, pValue);
                subjectFilter.Append(condition);
            }

            return subjectFilter;
        }
    }
}
