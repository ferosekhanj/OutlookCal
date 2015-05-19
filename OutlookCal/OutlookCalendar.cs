using System.Collections.Generic;
/*
 * Created by SharpDevelop.
 * User: ic003194
 * Date: 2/6/2012
 * Time: 5:12 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace OutlookCal
{
	using System;
	using Microsoft.Office.Interop.Outlook;
	
	/// <summary>
	/// Description of OutlookCalendar.
	/// </summary>
	public class OutlookCalendar
	{
		Application myApp;
		
		NameSpace myMapiNS;
		MAPIFolder myCalendar;
		
		public OutlookCalendar()
		{
			myApp 		= new ApplicationClass();
			myMapiNS 	= myApp.GetNamespace("MAPI");
			myCalendar 	= myMapiNS.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
		}
		
		public IList<ApptItem> QueryCalendarForTheDay(DateTime aToday)
		{
			List<ApptItem> anItems = new List<ApptItem>();
			DateTime aTomorrow = aToday.Date.AddDays(1);
			foreach( AppointmentItem anAppt in myCalendar.Items)
			{
                AppointmentItem currentApt = anAppt;
				DateTime aStart =  anAppt.Start;
                DateTime anEnd = anAppt.End;
                if (!anAppt.IsRecurring)
				{
					if( anEnd < aToday || anEnd > aTomorrow || aStart > aTomorrow)
					{
						continue;
					}
				}
				else if (!DoesOccur(anAppt,aToday,aTomorrow))
				{
					continue;	
				}
				
				if( anAppt.IsRecurring)
				{
					if( (currentApt = GetOccurence(anAppt,aToday)) == null)
					{
						continue;
					}
				}
				
				anItems.Add(
					new ApptItem() 
					{
                        Start = new DateTime(currentApt.Start.Ticks),
                        End = new DateTime(currentApt.End.Ticks),
                        Summary = currentApt.Subject,
                        Location = currentApt.Location						
					});
			}
			
			return anItems;
		}
		
		bool DoesOccur(AppointmentItem theAppt, DateTime theToday, DateTime theTomorrow)
		{
			RecurrencePattern anObjPattern = theAppt.GetRecurrencePattern();
			if( anObjPattern.NoEndDate || anObjPattern.PatternEndDate > theToday)
			{
				return DoesOccurToday(theAppt, anObjPattern, theToday);
			}
			return false;
		}
		
		bool DoesOccurToday(AppointmentItem theAppt, RecurrencePattern theObjPattern, DateTime theToday)
		{
			switch (theObjPattern.RecurrenceType) {
				case OlRecurrenceType.olRecursDaily:
					return true;
				case OlRecurrenceType.olRecursWeekly:
                    return DoesOccurInThisWeek(theAppt, theObjPattern, theToday);
                case OlRecurrenceType.olRecursMonthly:
                    return DoesOccurInThisMonth(theAppt, theObjPattern, theToday);
                default:
					return false;
			}
		}

        private bool DoesOccurInThisMonth(AppointmentItem theAppt, RecurrencePattern theObjPattern, DateTime theToday)
        {
            DateTime anOcc = new DateTime(theToday.Year, theToday.Month, theToday.Day,
                                         theAppt.Start.Hour, theAppt.Start.Minute, theAppt.Start.Second);

            return (theObjPattern.DayOfMonth == theToday.Day);
        }
		
		bool DoesOccurInThisWeek(AppointmentItem theAppt, RecurrencePattern theObjPattern, DateTime theToday)
		{
			DateTime anOcc = new DateTime(theToday.Year, theToday.Month, theToday.Day,
			                             theAppt.Start.Hour, theAppt.Start.Minute,theAppt.Start.Second);
			
			if ( ( (1 << ((int)theToday.DayOfWeek)) & ((int)theObjPattern.DayOfWeekMask) ) > 0)
			{
  			 	int offset = theAppt.Start.Day - theToday.Day;
		        TimeSpan timeSpan = ( anOcc - theAppt.Start ); 
		        return ( (timeSpan.TotalDays+offset / 7) % theObjPattern.Interval == 0 );
			}
			
			return false;
		}

        AppointmentItem GetOccurence(AppointmentItem theAppt, DateTime theToday)
		{
			AppointmentItem apptOccu = null;
			DateTime aDay = new DateTime(theToday.Year,theToday.Month,theToday.Day,
			                             theAppt.Start.Hour-1,theAppt.Start.Minute,0);
			for(int i = 0; i < 3; i++)
			{
				try
				{
					apptOccu = theAppt.GetRecurrencePattern().GetOccurrence(aDay);
					break;
				}
				catch(System.Exception e)
				{
					aDay =aDay.AddHours(1);
				}
			}
			return apptOccu;
		}
	}
	
	public class ApptItem
	{
		public DateTime Start{get;set;}
		public DateTime End{get;set;}
		public String Summary{get;set;}
		public String Location {get;set;}
	}
}
