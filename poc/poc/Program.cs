using System;
using Microsoft.Exchange.WebServices.Data;
namespace poc
{
    class Program
    {
        
        private ExchangeService Service
        {
            get
            {
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010);
                service.Credentials = new WebCredentials("email@email.com", "password");
                service.Url = new Uri("https://outlook.office365.com/ews/exchange.asmx");
                return service;
            }
        }

        private CalendarFolder FindDefaultCalendarFolder()
        {
            return CalendarFolder.Bind(Service, WellKnownFolderName.Calendar, new PropertySet());
        }

        private void LoadAppointments()
        {
            DateTime startDate = FirstDayOfWeek.GetFirstDayOfWeek(DateTime.Now);
            DateTime endDate = startDate.AddDays(7);

            CalendarFolder calendar = FindDefaultCalendarFolder();

            CalendarView cView = new CalendarView(startDate, endDate, 50);
            cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.Id);
            FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);

            foreach (Appointment appointment in appointments)
            {
                Console.WriteLine("Subject: " + appointment.Subject);
            }

        }
        static void Main(string[] args)
        {
            Program p = new Program();
            p.LoadAppointments();

            Console.ReadKey();
        }
    }
}
