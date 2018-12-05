using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

        /* Constructors */

        public Calendar(string pUsername, string pPassword, bool pShowDialog, bool pNewSession, string pNamespace, string pFolderID)
        {
            try
            {
                _OutlookApplication = new Outlook.Application();
                _OutlookNameSpace = _OutlookApplication.GetNamespace(pNamespace);
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
                _OutlookNameSpace.Logon(username, password, pShowDialog, pNewSession);

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

        public Calendar(string pUsername) : this(pUsername, null) { }

        public Calendar() : this(null) { }

        /* Public methods */

        public void Disconnect()
        {
            _OutlookNameSpace.Logoff();

            _OutlookApplication = null;
            _OutlookCalendar = null;
            _OutlookNameSpace = null;
            _CalendarItems = null;
        }


        //Returning every appointment within a given interval 
        public List<Outlook.AppointmentItem> FindAppointmentsInRange(DateTime pStart, DateTime pEnd)
        {
            List<Outlook.AppointmentItem> appts = new List<Outlook.AppointmentItem>();
            foreach (Outlook.AppointmentItem appt in RestrictByInterval(pStart, pEnd))
            {
                appts.Add(appt);
            }
            return appts;
        }

        //Returns every appointment with a given subject within a given interval
        public List<Outlook.AppointmentItem> FindAppointmentsBySubject(string pSubject, DateTime pStart, DateTime pEnd, bool pIsSubstring)
        {
            return RestrictBySubject(RestrictByInterval(pStart, pEnd), pSubject, pIsSubstring);
        }

        public List<Outlook.AppointmentItem> FindAppointmentsBySubject(string pSubject, DateTime pStart, DateTime pEnd)
        {
            return FindAppointmentsBySubject(pSubject, pStart, pEnd, false);
        }

        //Returns every appointment with one of the given subjects within a given interval
        public List<Outlook.AppointmentItem> FindAppointmentsBySubject(string[] pSubjects, DateTime pStart, DateTime pEnd, bool pIsSubstring)
        {
            Outlook.Items apptsInInterval = RestrictByInterval(pStart, pEnd);
            List<Outlook.AppointmentItem> appts = new List<Outlook.AppointmentItem>();
            foreach (string subject in pSubjects)
            {
                appts.Concat(RestrictBySubject(RestrictByInterval(pStart, pEnd), subject, pIsSubstring));
            }
            return appts;
        }

        public List<Outlook.AppointmentItem> FindAppointmentsBySubject(string[] pSubjects, DateTime pStart, DateTime pEnd)
        {
            return FindAppointmentsByLocation(pSubjects, pStart, pEnd, false);
        }

        //Returns every appointment in a given location within a given interval 
        public List<Outlook.AppointmentItem> FindAppointmentsByLocation(string pLocation, DateTime pStart, DateTime pEnd, bool pIsSubstring)
        {
            return RestrictByLocation(RestrictByInterval(pStart, pEnd), pLocation, pIsSubstring);
        }

        public List<Outlook.AppointmentItem> FindAppointmentsByLocation(string pLocation, DateTime pStart, DateTime pEnd)
        {
            return FindAppointmentsByLocation(pLocation, pStart, pEnd, false);
        }

        //Returns every appointment with one of the given locations within a given interval
        public List<Outlook.AppointmentItem> FindAppointmentsByLocation(string[] pLocations, DateTime pStart, DateTime pEnd, bool pIsSubstring)
        {
            Outlook.Items apptsInInterval = RestrictByInterval(pStart, pEnd);
            List<Outlook.AppointmentItem> appts = new List<Outlook.AppointmentItem>();
            foreach (string location in pLocations)
            {
                appts.Concat(RestrictByLocation(RestrictByInterval(pStart, pEnd), location, pIsSubstring));
            }
            return appts;
        }

        public List<Outlook.AppointmentItem> FindAppointmentsByLocation(string[] pLocations, DateTime pStart, DateTime pEnd)
        {
            return FindAppointmentsByLocation(pLocations, pStart, pEnd, false);
        }

        //--------------------------------------------------------------------------------------------------------------
        /* Methods for class internal use */
        //--------------------------------------------------------------------------------------------------------------

        private Outlook.Items RestrictByInterval(DateTime pStart, DateTime pEnd)
        {
            string filter = String.Format("[Start] >= '{0}' AND [End] < '{1}'",
                pStart.ToString("g"), pEnd.AddDays(1).ToString("g"));

            Outlook.Items callItems = _CalendarItems;
            callItems.IncludeRecurrences = true;
            callItems.Sort("[Start]", Type.Missing);
            Outlook.Items restrictedItems = callItems.Restrict(filter);
            if (restrictedItems.Count > 0)
            {
                return restrictedItems;
            }
            else
            {
                return null;
            }
        }

        //Restrict a list of appointments to the ones with a given subject
        private static List<Outlook.AppointmentItem> RestrictBySubject(Outlook.Items pAppointments, string pSubject, bool pIsSubstring)
        {

            string filter = BuildFilterString("Subject", pSubject, pIsSubstring, "urn:schema:httpmail:subject").ToString();
            return Restrict(pAppointments, filter);
        }

        //Restrict a list of appointments to the ones with  a given location
        private static List<Outlook.AppointmentItem> RestrictByLocation(Outlook.Items pAppointments, string pLocation, bool pIsSubstring)
        { 
            string filter = BuildFilterString("Location", pLocation, pIsSubstring, "urn:schema:calendar:location").ToString();
            return Restrict(pAppointments, filter);   
        }

        //Restricts a List of Outlook.Items to the ones matching a given filter an returns them as a list of Outlook.AppointmentsItems
        private static List<Outlook.AppointmentItem> Restrict(Outlook.Items pItems, string pFilter)
        {
            List<Outlook.AppointmentItem> appts = new List<Outlook.AppointmentItem>();
            Outlook.AppointmentItem appointment = pItems.Find(pFilter);

            while (appointment != null)
            {
                appts.Add(appointment);
                appointment = pItems.FindNext() as Outlook.AppointmentItem;
            }

            if (appts.Count > 0)
            {
                return appts;
            }
            else
            {
                return null;
            }
        }

        //Builds a filter-string that can be used to search for appointments matching this filter
        private static StringBuilder BuildFilterString(string pKeyword, string pValue, bool pIsSubstring, string pCriteria)
        {
            StringBuilder subjectFilter = new StringBuilder();
            if (pIsSubstring == true)
            {
                subjectFilter.Append("SQL=");
                subjectFilter.Append(pCriteria);
                string valueCondition = String.Format(" like '%{0}%'", pValue);
            }
            else
            {
                string condition = String.Format("[{0}] = {1}", pKeyword, pValue);
            }

            return subjectFilter;
        }
    }
}
