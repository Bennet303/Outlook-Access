using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Reflection;

namespace Outlook_Access
{ 
    public class Calendar
    {
        private Outlook.Application _OutlookApplication;
        private Outlook.NameSpace _OutlookNameSpace;
        private Outlook.MAPIFolder _OutlookCalendar;
        private Outlook.Items _CalendarItems;

        public Calendar(string pUsername, string pPassword, bool pShowDialog, bool pNewSession, string pNamespace, string pFolderID)
        {
            try
            {
                _OutlookApplication = new Outlook.Application();
                _OutlookNameSpace = _OutlookApplication.GetNamespace(pNamespace);

                if (pUsername != null && pPassword != null)
                {
                    _OutlookNameSpace.Logon(pUsername, pPassword);
                }
                else
                {
                    _OutlookNameSpace.Logon(Missing.Value, Missing.Value);
                }

                if (pFolderID == null)
                {
                    _OutlookCalendar = _OutlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                }
                else
                {
                    _OutlookCalendar = _OutlookNameSpace.GetFolderFromID(pFolderID);
                }

                _CalendarItems = _OutlookCalendar.Items;
            }
            catch (Exception e)
            {
                Console.WriteLine($"{e} Exception caught");
            }
        }

        public Calendar(string pUsername, string pPassword, bool pShowDialog, bool pNewSession, string pFolderID) : this(pUsername, pPassword, pShowDialog, pNewSession, "mapi", pFolderID) { }

        public Calendar(string pUsername, string pPassword, bool pShowDialog, bool pNewSession) : this(pUsername, pPassword, pShowDialog, pNewSession, "mapi", null) { }

        public Calendar(string pUsername, string pPassword) : this(pUsername, pPassword, false, true) { }

        public Calendar(string pFolderID) : this(null, null, false, true, pFolderID) { }

        public Calendar() : this(null) { }



        //Returning every appointment within a given interval 
        public Outlook.Items FindAppointmentsInRange(DateTime pStart, DateTime pEnd)
        {
            string filter = "[Start] >=\'" + pStart.ToString("g") + "' AND [END] <= '" + pEnd.ToString("g") + "'";

            Outlook.Items restrictedItems = _OutlookCalendar.Items.Restrict(filter);
            restrictedItems.Sort("[Start]", Type.Missing);
            restrictedItems.IncludeRecurrences = true;
            if (restrictedItems.Count > 0)
            {
                return restrictedItems;
            }
            else
            {
                return null;
            }
        }


    }
}
