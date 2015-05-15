/*
 * Created by SharpDevelop.
 * User: ic003194
 * Date: 2/7/2012
 * Time: 11:15 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Diagnostics;
using System.IO;
using System.Xml;

using DotNetOpenAuth.Messaging;
using DotNetOpenAuth.OAuth2;
using Google.Apis.Authentication.OAuth2;
using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Util;
using Microsoft.Office.Interop.Outlook;

namespace OutlookCal
{
	
	
	/// <summary>
	/// Description of GoogleCalendar.
	/// </summary>
	public class GoogleCalendar
	{
		CalendarService myCal;
        String myGmail;
		public GoogleCalendar(String theGmail)
		{
			Login();
            myGmail = theGmail;
		}
		
		private void Login()
		{
			NativeApplicationClient aProvider = new NativeApplicationClient(GoogleAuthenticationServer.Description);
			aProvider.ClientIdentifier = MyGoogleSecret.ClientIdentifier;
			aProvider.ClientSecret = MyGoogleSecret.ClientSecret;
			var anAuth = new OAuth2Authenticator<NativeApplicationClient>(aProvider, GetAuthentication);
			myCal = new CalendarService(anAuth);
		}
		
		private IAuthorizationState GetAuthentication(NativeApplicationClient theClient)
		{
			// Try to fetch from the cache
			IAuthorizationState aState = new AuthorizationState(new[] { CalendarService.Scopes.Calendar.GetStringValue()});
			string aRefreshToken = LoadCredentialsFromCache();
			if( aRefreshToken != null)
			{
				try
				{
					aState.RefreshToken = aRefreshToken;
					theClient.RefreshToken(aState);
					return aState;
				}
				catch(ProtocolException pe)
				{
					Console.WriteLine(pe.ToStringDescriptive());
					Console.WriteLine("Using existing token failed!!");
				}
			}
			aState.Callback = new Uri(NativeApplicationClient.OutOfBandCallbackUrl);
		    // If it fails then ask for the autorization again from google
		    Uri anAuthUri = theClient.RequestUserAuthorization(aState);
		
		    // Request authorization from the user (by opening a browser window):
		    Process.Start(anAuthUri.ToString());
		    Console.Write("Please enter the authorization Code: ");
		    string anAuthCode = Console.ReadLine();
		    Console.WriteLine();
		    // Retrieve the access token by using the authorization code:
		    IAuthorizationState aResultState = theClient.ProcessUserAuthorization(anAuthCode, aState);
		    CacheCredentials(aResultState.RefreshToken);
		    return aResultState;
		}
		
		private string LoadCredentialsFromCache()
		{
			if( File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)+"\\OutlookCal\\goog.auth"))
			{
				FileStream fs = new FileStream(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)+"\\OutlookCal\\goog.auth",FileMode.Open);
				StreamReader sr = new StreamReader(fs);
				string aRefreshToken = sr.ReadLine();
				sr.Close();
				fs.Close();
				return aRefreshToken;
			}
			return null;
		}
		
		private void CacheCredentials(string theRefreshToken)
		{
			if( !Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)+"\\OutlookCal"))
			{
				Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)+"\\OutlookCal");
			}
			if( File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)+"\\OutlookCal\\goog.auth"))
			{
				File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)+"\\OutlookCal\\goog.auth");
			}
			
			FileStream fs = new FileStream(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)+"\\OutlookCal\\goog.auth",FileMode.CreateNew);
			StreamWriter sw = new StreamWriter(fs);
			sw.WriteLine(theRefreshToken);
			sw.Close();
			fs.Close();
		}
		
		public void ListTodaysAppt()
		{
            var items = myCal.Events.List(myGmail).Fetch();
		}
		
		public void AddEvent(DateTime theStart, DateTime theEnd, string theSummary, string theLocation)
		{
			EventDateTime aStart = new EventDateTime(){DateTime = XmlConvert.ToString(theStart,XmlDateTimeSerializationMode.Local)};
			EventDateTime aEnd = new EventDateTime(){DateTime = XmlConvert.ToString(theEnd,XmlDateTimeSerializationMode.Local)};
			Event aNewEvent = myCal.Events.Insert(
				new Event
				{ 
					Start = aStart, 
					End = aEnd, 
					Summary = theSummary,
					Location = theLocation
				},
                myGmail).Fetch();
		}
	}
}
