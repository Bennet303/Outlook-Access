using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outlook_Access
{
    
    public class Contacts : AccessClass
    {
        //--------------------------------------------------------------------------------------------------------------
        /* Constructors*/
        //--------------------------------------------------------------------------------------------------------------
        

        public Contacts(string pUsername, string pPassword, bool pShowDialog, bool pNewSession, string pNamespace, string pFolderID)
        {
            Connect(pUsername, pPassword, pShowDialog, pNewSession, pNamespace, pFolderID);
            if (pFolderID == null)
            {
                OutlookFolder = OutlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            }
            OutlookFolderItems = OutlookFolder.Items;
        }
        public Contacts(string pUsername, string pPassword, bool pShowDialog, bool pNewSession, string pFolderID) : this(pUsername, pPassword, pShowDialog, pNewSession, "mapi", pFolderID) { }

        public Contacts(string pUsername, string pPassword, bool pShowDialog, bool pNewSession) : this(pUsername, pPassword, pShowDialog, pNewSession, "mapi", null) { }

        public Contacts(string pUsername, string pPassword) : this(pUsername, pPassword, false, true) { }

        public Contacts(string pUsername) : this(pUsername, null) { }

        public Contacts() : this(null) { }



        //--------------------------------------------------------------------------------------------------------------
        /* Reading*/
        //--------------------------------------------------------------------------------------------------------------

        public List<Outlook.ContactItem> FindContactsByFirstName(string pFirstName, bool pSubstring)
        {
            string filter = BuildFilterString("FirstName", pFirstName, pSubstring, "urn:schemas:contacts:givenName").ToString();
            return Restrict(OutlookFolderItems, filter);
        }

        public List<Outlook.ContactItem> FindContactsByFirstName(string pFirstName)
        {
            return FindContactsByFirstName(pFirstName, false);
        }

        public List<Outlook.ContactItem> FindContactsByLastName(string pLastName, bool pSubstring)
        {
            string filter = BuildFilterString("LastName", pLastName, pSubstring, "urn:schemas:contacts:sn").ToString();
            return Restrict(OutlookFolderItems, filter);
        }

        public List<Outlook.ContactItem> FindContactsByLastName(string pLastName)
        {
            return FindContactsByLastName(pLastName, false);
        }

        public List<Outlook.ContactItem> FindContactsByFullName(string pFirstName, string pLastName, bool pSubstring)
        {
            //TODO: Add substring search
            StringBuilder filter = new StringBuilder();
            filter.Append(BuildFilterString("FirstName", pFirstName, pSubstring, "").ToString());
            filter.Append(" and ");
            filter.Append(BuildFilterString("LastName", pLastName, pSubstring, "").ToString());

            return Restrict(OutlookFolderItems, filter.ToString()) as List<Outlook.ContactItem>;
        }

        public List<Outlook.ContactItem> FindContactsByFullName(string pFirstName, string pLastName)
        {
            return FindContactsByFullName(pFirstName, pLastName, false) as List<Outlook.ContactItem>;
        }

        



        //--------------------------------------------------------------------------------------------------------------
        /* Writing */
        //--------------------------------------------------------------------------------------------------------------



        //--------------------------------------------------------------------------------------------------------------
        /* Methods for class internal use */
        //--------------------------------------------------------------------------------------------------------------

        private static List<Outlook.ContactItem> Restrict(Outlook.Items pContacts, string pFilter)
        {
            List<Outlook.ContactItem> contacts = new List<Outlook.ContactItem>();
            Outlook.Items contactsInFolder = pContacts.Restrict(pFilter);
            foreach (Outlook.ContactItem contact in contactsInFolder)
            {
                contacts.Add(contact);
            }
            return contacts as List<Outlook.ContactItem>;
        }
    }
}
