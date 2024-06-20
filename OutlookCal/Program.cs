/*
 * Created by SharpDevelop.
 * User: ic003194
 * Date: 2/6/2012
 * Time: 4:34 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;

namespace OutlookCal
{
	class Program
	{
		public static void Main(string[] args)
		{
			DateTime aToday = DateTime.Now.Date;
            if (args.Length == 0)
            {
                Console.WriteLine("OutlookCal.exe <username@gmail.com>");
                return;
            }
            Console.WriteLine($"{aToday.ToShortDateString()}\r\n==========");
            OutlookCalendar aCal = new OutlookCalendar();
			
			IList<ApptItem> anItemsToArchive = aCal.QueryCalendarForTheDay(aToday);

			if( anItemsToArchive.Count == 0 )
			{
				Console.WriteLine("Nothing to archive!!");
				return;
			}

			GoogleCalendar aCloudCal = new GoogleCalendar(args[0]);

			foreach (ApptItem anAppt in anItemsToArchive)
			{
				Console.WriteLine("{0} {1} {2} {3} {4}", aToday.ToShortDateString(), anAppt.Start.ToShortTimeString(), anAppt.End.ToShortTimeString(), anAppt.Summary, anAppt.Location);
				aCloudCal.AddEvent(anAppt.Start, anAppt.End, anAppt.Summary, anAppt.Location);
			}

			Console.Write("\r\nPress any key to continue . . . ");
			Console.ReadKey(true);
		}
	}
}
