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

        /* Constructors */

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
        public Outlook.Items FindAppointmentsInRange(DateTime pStart, DateTime pEnd)
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

        //Returns every appointment with a given subject within a given interval
        public Outlook.Items FindAppointmentsBySubject(string pSubject, DateTime pStart, DateTime pEnd)
        {
            return RestrictBySubject(FindAppointmentsInRange(pStart, pEnd), pSubject);
        }

        //Returns every appointment with one of the given subjects within a given interval
        public Outlook.Items FindAppointmentsBySubject(string[] pSubjects, DateTime pStart, DateTime pEnd)
        {
            Outlook.Items apptsInInterval = FindAppointmentsInRange(pStart, pEnd);
            Outlook.Items appts = new Outlook.Items();
            foreach (string subject in pSubjects)
            {
                foreach (Outlook.AppointmentItem item in RestrictBySubject(apptsInInterval, subject))
                {
                    appts.Add(item);
                }
            }
            appts.Sort("[Start]", Type.Missing);
            return appts;
        }

        //Returns every appointment in a given location within a given interval 
        public Outlook.Items FindAppointmentsByLocation(string pLocation, DateTime pStart, DateTime pEnd)
        {
            return RestrictByLocation(FindAppointmentsInRange(pStart, pEnd), pLocation);
        }

        //Returns every appointment with one of the given locations within a given interval
        public Outlook.Items FindAppointmentsByLocation(string[] pLocations, DateTime pStart, DateTime pEnd)
        {
            Outlook.Items apptsInInterval = FindAppointmentsInRange(pStart, pEnd);
            Outlook.Items appts = new Outlook.Items();
            foreach (string location in pLocations)
            {
                foreach (Outlook.AppointmentItem item in RestrictByLocation(apptsInInterval, location))
                {
                    appts.Add(item);
                }
            }
            appts.Sort("[Start]", Type.Missing);
            return appts;
        }

        /* Methods for class internal use */

        //Restrict a list of appointments to the ones with a given subject
        private Outlook.Items RestrictBySubject(Outlook.Items pAppointments, string pSubject)
        {
            StringBuilder subjectFilter = new StringBuilder();
            subjectFilter.Append("@SQL=" + "\"" + "urn:schemas:httpmail:subject" + "\"");

            if (_OutlookApplication.Session.DefaultStore.IsInstantSearchEnabled)
            {
                subjectFilter.Append(@" ci_startswith '{pSubject}'");
            }
            else
            {
                subjectFilter.Append(@" like '%{pSubject}%'");
            }

            Outlook.Items appts = pAppointments;
            appts.Restrict(subjectFilter.ToString());
            if (appts.Count > 0)
            {
                return appts;
            }
            else
            {
                return null;
            }
        }

        //Restrict a list of appointments to the ones with  a given location
        private Outlook.Items RestrictByLocation(Outlook.Items pAppointments, string pLocation)
        {
            StringBuilder locationFilter = new StringBuilder();
            locationFilter.Append("@SQL=" + "\"" + "urn:schemas:httpmail:location" + "\"");

            if (_OutlookApplication.Session.DefaultStore.IsInstantSearchEnabled)
            {
                locationFilter.Append(@" ci_startswith '{pLocation}'");
            }
            else
            {
                locationFilter.Append(@" like '%{pLocation}%'");
            }

            Outlook.Items appts = pAppointments;
            appts.Restrict(locationFilter.ToString());
            if (appts.Count > 0)
            {
                return appts;
            }
            else
            {
                return null;
            }
        }
    }
}
