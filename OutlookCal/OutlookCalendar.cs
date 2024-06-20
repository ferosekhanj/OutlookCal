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
            myApp       = new Microsoft.Office.Interop.Outlook.Application();
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
                DateTime aStart = anAppt.Start;
				DateTime anEnd  = anAppt.End;
				String ConversationTopic = anAppt.ConversationTopic??"None";
                RecurrencePattern anObjPattern = (anAppt.IsRecurring)?anAppt.GetRecurrencePattern():null;

                if ( !anAppt.IsRecurring)
				{
					if( anEnd <= aToday || anEnd > aTomorrow || aStart > aTomorrow)
					{
						continue;
					}
				}
				else if (!DoesOccur(anAppt,anObjPattern,aToday))
				{
                    if (!DoesAnExceptionOccurToday(anAppt,anObjPattern,aToday))
					{
						continue;
					}
				}
		
				if(anAppt.IsRecurring)
                {
                    if((currentApt = GetOccurence(anAppt,anObjPattern,aToday))==null)
					{
						continue;
					}
                }

				anItems.Add(
					new ApptItem() 
					{ 
						Start = new DateTime(aToday.Year,aToday.Month,aToday.Day,currentApt.Start.Hour,currentApt.Start.Minute,currentApt.Start.Second), 
						End = new DateTime(aToday.Year,aToday.Month,aToday.Day,currentApt.End.Hour,currentApt.End.Minute,currentApt.End.Second),
						Summary = currentApt.Subject,
						Location = currentApt.Location						
					});
			}
			anItems.Sort((x,y) => x.Start.CompareTo(y.Start));
			return anItems;
		}
		
		bool DoesOccur(AppointmentItem theAppt, RecurrencePattern theObjPattern, DateTime theToday)
		{
			if(theObjPattern.NoEndDate || theObjPattern.PatternEndDate > theToday)
			{
				return DoesOccurToday(theAppt, theObjPattern, theToday);
			}
			return false;
		}
		
		bool DoesOccurToday(AppointmentItem theAppt, RecurrencePattern theObjPattern, DateTime theToday)
		{
			switch (theObjPattern.RecurrenceType) {
				case OlRecurrenceType.olRecursDaily:
					return true;
				case OlRecurrenceType.olRecursWeekly:
				case OlRecurrenceType.olRecursMonthly:
                case OlRecurrenceType.olRecursMonthNth:
					return DoesOccurInThisWeek(theAppt,theObjPattern,theToday);
				default:
					throw new System.Exception("Invalid value for OlRecurrenceType");
			}
		}
		
		bool DoesAnExceptionOccurToday(AppointmentItem theAppt, RecurrencePattern theObjPattern, DateTime theToday)
        {
            foreach (Microsoft.Office.Interop.Outlook.Exception anExcept in theObjPattern.Exceptions)
            {
				if (anExcept.Deleted)
				{
					continue;
				}
                if( anExcept.AppointmentItem.Start.Date == theToday.Date)
                {
                    return true;
                }
            }
            return false;
        }

		bool DoesOccurInThisWeek(AppointmentItem theAppt, RecurrencePattern theObjPattern, DateTime theToday)
		{
			DateTime anOcc = new DateTime(theToday.Year, theToday.Month, theToday.Day,
			                             theAppt.Start.Hour, theAppt.Start.Minute,theAppt.Start.Second);
			
			if ( ( (1 << ((int)theToday.DayOfWeek)) & ((int)theObjPattern.DayOfWeekMask) ) > 0)
			{
				int offset = theAppt.Start.DayOfWeek - theToday.DayOfWeek;
		        TimeSpan timeSpan = ( anOcc - theAppt.Start ); 
				timeSpan = timeSpan.Add(new TimeSpan(offset,0,0,0));
		        return ( ((long)timeSpan.TotalDays) % (theObjPattern.Interval*7) == 0 );
			}
			
			return false;
		}
		
        AppointmentItem GetOccurence(AppointmentItem theAppt, RecurrencePattern theObjPattern, DateTime theToday)
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
			// if not found search in the exceptions
			if(apptOccu == null)
			{
                foreach (Microsoft.Office.Interop.Outlook.Exception anExcept in theObjPattern.Exceptions)
                {
                    if (anExcept.Deleted)
                    {
                        continue;
                    }
                    if (anExcept.AppointmentItem.Start.Date == theToday.Date)
                    {
                        apptOccu = anExcept.AppointmentItem;
						break;
                    }
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
