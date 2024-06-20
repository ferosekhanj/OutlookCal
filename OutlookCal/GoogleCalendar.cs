/*
 * Created by SharpDevelop.
 * User: ic003194
 * Date: 2/7/2012
 * Time: 11:15 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.IO;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;

namespace OutlookCal
{
	/// <summary>
	/// Description of GoogleCalendar.
	/// </summary>
	public class GoogleCalendar
	{
		CalendarService myCal;
		string myCalendarId;

		public GoogleCalendar(string theCalendarId)
		{
            myCalendarId = theCalendarId;
			Login();
		}

		private void Login()
		{
			UserCredential credential;
			using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
			{
				credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
					GoogleClientSecrets.FromStream(stream).Secrets,
					new[] { CalendarService.Scope.Calendar },
					"user",
					System.Threading.CancellationToken.None).Result;
			}
			myCal = new CalendarService(new BaseClientService.Initializer()
			{
				HttpClientInitializer = credential,
				ApplicationName = "OutlookCal"
			});
		}

		public void ListTodaysAppt()
		{
			var result = myCal.Events.List(myCalendarId).Execute();

			foreach(var item in result.Items)
			{
				Console.WriteLine($"{item.Id} {item.Summary}");
			}
		}
		
		public void AddEvent(DateTime theStart, DateTime theEnd, string theSummary, string theLocation)
		{
			EventDateTime aStart = new EventDateTime() { DateTime = theStart };
			EventDateTime aEnd = new EventDateTime() { DateTime = theEnd };
			Event anEvent = new Event()
            {
                Start = aStart,
                End = aEnd,
                Summary = theSummary,
                Location = theLocation
            };
			Event aNewEvent = myCal.Events.Insert(anEvent,myCalendarId).Execute();
		}
	}
}
