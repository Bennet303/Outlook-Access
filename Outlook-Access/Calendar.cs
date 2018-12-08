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

        /*--------------------------------------------------------------------------------------------
         * Reading
         -------------------------------------------------------------------------------------------*/

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

        /*--------------------------------------------------------------------------------------------
         * Writing
         -------------------------------------------------------------------------------------------*/
        public void AddAppointment(string pSubject, DateTime pStart, DateTime pEnd, string pLocation, bool pAllDay, string pBody, string pCategory)
        {
            InternalAddAppointment(pSubject, pStart, pEnd, pLocation, pAllDay, pBody, pCategory);
        }

        public void AddAppointment(string pSubject, DateTime pStart, double pLength, string pLocation, bool pAllDay, string pBody, string pCategory)
        {
            AddAppointment(pSubject, pStart, pStart.AddHours(pLength), pLocation, pAllDay, pBody, pCategory);
        }

        public void AddAppointment(string pSubject, DateTime pStart, DateTime pEnd, string pLocation, bool pAllDay, string pBody)
        {
            AddAppointment(pSubject, pStart, pEnd, pLocation, pAllDay, pBody, null);
        }

        public void AddAppointment(string pSubject, DateTime pStart, DateTime pEnd, string pLocation, bool pAllDay)
        {
            AddAppointment(pSubject, pStart, pEnd, pLocation, pAllDay, null);
        }

        public void AddAppointment(string pSubject, DateTime pStart, DateTime pEnd, string pLocation, string pBody)
        {
            AddAppointment(pSubject, pStart, pEnd, pLocation, false, pBody);
        }

        public void AddAppointment(string pSubject, DateTime pStart, double pLength, string pLocation)
        {
            AddAppointment(pSubject, pStart, pLength, pLocation, false, null, null);
        }

        public void AddAppointment(string pSubject, DateTime pStart, DateTime pEnd)
        {
            AddAppointment(pSubject, pStart, pEnd, null, false);
        }

        public void AddAppointment(string pSubject, DateTime pStart)
        {
            AddAppointment(pSubject, pStart, pStart.AddHours(1.5));
        }

        public Outlook.AppointmentItem AddReccuringAppointment(string pSubject, DateTime pStart, DateTime pEnd, string pLocation, bool pAllDay, string pBody, string pCategory)
        {
            return InternalAddAppointment(pSubject, pStart, pEnd, pLocation, pAllDay, pBody, pCategory);
        }

        public void EditRecurringItem(Outlook.AppointmentItem pRecurringAppointment, DateTime pPatternStart, DateTime pPatternEnd,
            int pOccurrences, string pRecurrenceType, Enum pDaysOfWeek)
        {
            try
            {
                Outlook.RecurrencePattern recurrencePattern = pRecurringAppointment.GetRecurrencePattern();
                recurrencePattern.PatternStartDate = pPatternStart.Date;

                if (pPatternEnd != null)
                {
                    recurrencePattern.PatternEndDate = pPatternEnd;
                }
                else
                {
                    recurrencePattern.NoEndDate = true;
                }

                if (pOccurrences != 0)
                {
                    recurrencePattern.Occurrences = pOccurrences;
                }

                Outlook.OlRecurrenceType recurrenceType;
                switch (pRecurrenceType.ToLower())
                {
                    case "daily": recurrenceType = Outlook.OlRecurrenceType.olRecursDaily; break;
                    case "weekly": recurrenceType = Outlook.OlRecurrenceType.olRecursWeekly; break;
                    case "monthly": recurrenceType = Outlook.OlRecurrenceType.olRecursMonthly; break;
                    case "yearly": recurrenceType = Outlook.OlRecurrenceType.olRecursYearly; break;
                    default: throw new ArgumentNullException("No viable RecurrenceType selected");
                }
                recurrencePattern.RecurrenceType = recurrenceType;

                recurrencePattern.DayOfWeekMask = (Outlook.OlDaysOfWeek)pDaysOfWeek;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        //--------------------------------------------------------------------------------------------------------------
        /* Methods for class internal use */
        //--------------------------------------------------------------------------------------------------------------


        private Outlook.AppointmentItem InternalAddAppointment(string pSubject, DateTime pStart, DateTime pEnd, string pLocation, bool pAllDay, string pBody, string pCategory)
        {
            try
            {
                Outlook.AppointmentItem appointment = (Outlook.AppointmentItem)
                    _OutlookApplication.CreateItem(Outlook.OlItemType.olAppointmentItem);
                appointment.Subject = pSubject;
                appointment.Start = pStart;
                appointment.End = pEnd;
                appointment.AllDayEvent = pAllDay;
                if (pLocation != null)
                {
                    appointment.Location = pLocation;
                }
                if (pBody != null)
                {
                    appointment.Body = pBody;
                }
                if (pCategory != null)
                {
                    appointment.Categories = pCategory;
                }

                SaveAndDisplayAppointment(appointment);
                return appointment;
            }
            catch (Exception e)
            {
                Console.WriteLine($"{e} Exception caught");
                return null;
            }
        }

        private static void SaveAndDisplayAppointment(Outlook.AppointmentItem pAppt)
        {
            pAppt.Save();
            pAppt.Display(true);
        }
        
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
