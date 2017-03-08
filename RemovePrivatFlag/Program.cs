/*
 * 
 * User: Torsten Schlopsnies
 * Date: 24.02.2017
 * Time: 10:45
 * 
 *
 */
using System;
using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;

// Configure log4net using the .config file
[assembly: log4net.Config.XmlConfigurator(Watch=true)]
// This will cause log4net to look for a configuration file
// called ConsoleApp.exe.config in the application base
// directory (i.e. the directory containing ConsoleApp.exe)

namespace RemovePrivateFlag {
	class Program 	{
		private static FindFoldersResults findFolders;
		private static FindItemsResults<Item> findResults;
		private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
		
		public static void Main(string[] args) 		{
			log.Info("Application started");
			log.Debug("Parsing arguments");
			if (args.Length > 0) {		
				// getting all arguments from the command line
				var arguments = new UtilityArguments(args);
				if (arguments.Help)	{
					DisplayHelp();
					Environment.Exit(0);
				}
				string Postfach = arguments.Mailbox;
				if (Postfach == null) {
					if (log.IsWarnEnabled) {
						log.Warn("No mailbox is given. Use -help to refer to the usage.");
					}
					else {
						Console.WriteLine("No mailbox is given. Use -help to refer to the usage.");
					}
					DisplayHelp();
					Environment.Exit(1);
				}
				if (Postfach.Length == 0) {
					if (log.IsWarnEnabled) {
						log.Warn("No mailbox is given. Use -help to refer to the usage.");
					}
					else {
						Console.WriteLine("No mailbox is given. Use -help to refer to the usage.");
					}
					DisplayHelp();
					Environment.Exit(1);
				}
				// create the service
				ExchangeService ExService = ConnectToExchange(Postfach);
				if (log.IsInfoEnabled) log.Info("Service created.");
				// find all folders (under MsgFolderRoot)
				FindFoldersResults FolderList =  Folders(ExService);
				// check if we need to remove items from the list because we want to filter it (folderpath)
				string FolderName = arguments.Foldername;
				if (FolderName != null) {
					if (FolderName.Length > 0) {
						if (log.IsInfoEnabled)
							log.Info("Filter the folder list to apply filter.");
						for (int i = FolderList.Folders.Count - 1; i >= 0; i--) {
							if (log.IsDebugEnabled)
								log.Debug(string.Format("Processing folder for filtering: {0}", FolderList.Folders[i].DisplayName));
							try {
								string FolderPath;
								FolderPath = GetFolderPath(ExService, FolderList.Folders[i].Id);
								if (log.IsDebugEnabled)
									log.Debug(string.Format("Folderpath is: {0}", FolderPath));
								if (!(FolderPath.Contains(FolderName))) {
									log.Debug(string.Format("The folder: {0} does not match with the filter: {1}", FolderPath, FolderName));
									FolderList.Folders.RemoveAt(i);
								}
							} catch {
								Environment.Exit(2);
							}
						}
					}
				}
				// now try to find all items that are marked as "private"
				for(int i = FolderList.Folders.Count - 1; i >= 0; i--) {
					if (log.IsInfoEnabled) log.Info(string.Format("Processing folder {0}",GetFolderPath(ExService,FolderList.Folders[i].Id)));
					if (log.IsDebugEnabled) log.Debug(string.Format("ID: {0}", FolderList.Folders[i].Id));
					FindItemsResults<Item> Results = PrivateItems(FolderList.Folders[i]);
					foreach (var Result in Results) {
						if (Result is EmailMessage) {
							if (log.IsInfoEnabled) {
								log.Info(string.Format("Found private element. Folder: {0}",GetFolderPath(ExService,FolderList.Folders[i].Id)));
								log.Info(string.Format("Subject: {0}",Result.Subject));
								log.Debug(string.Format("ID of the item: {0}", Result.Id));
							}
							else {
								Console.WriteLine("Found private element. Folder: {0}",GetFolderPath(ExService,FolderList.Folders[i].Id));
								Console.WriteLine("Subject: {0}",Result.Subject);
							}
							if (!(arguments.noConfirmation)) {
								if (!(arguments.LogOnly)) {
									Console.WriteLine(string.Format("Change to normal? (y/N) (Folder: {0} - Subject {1})",GetFolderPath(ExService,FolderList.Folders[i].Id),Result.Subject));
									string Question = Console.ReadLine();
									if (Question == "y" || Question == "Y") {
										log.Info("Change the item? Answwer: Yes.");
										ChangeItem(Result);
									}
								}
							}
							else {
								if (!(arguments.LogOnly)) {
									if (log.IsInfoEnabled) {
										log.Info("Changing item without confirmation because -noconfirmation is true.");
									}
									else {
										Console.WriteLine("Changing item without confirmation because -noconfirmation is true.");
									}
									ChangeItem(Result);	
								}
							}							    
						}
					}
				}
			}
			else { 
				DisplayHelp();
				Environment.Exit(1);
			}
		}
		
		public static void DisplayHelp() {
			Console.WriteLine("Usage:");
			Console.WriteLine("RemovePrivateFlag.exe -mailbox \"user@example.com\" [-logonly true] [-foldername \"Inbox\" [-noconfirmation true]");
		}
		
		public static void ChangeItem(Item Message)	{
			// do we have the extended properties?
			if (Message.ExtendedProperties.Count > 0) {
				try	{
					var extendedPropertyDefinition = new ExtendedPropertyDefinition(0x36, MapiPropertyType.Integer);
					int extendedPropertyindex = 0;
										
					foreach (var extendedProperty  in Message.ExtendedProperties) {
									
						if (extendedProperty.PropertyDefinition == extendedPropertyDefinition) {
							if (log.IsInfoEnabled) {
								log.Info(string.Format("Try to remove private flag from message: {0}", Message.Subject));
							}
							else {
								Console.WriteLine("Try to remove private flag from message: {0}", Message.Subject);
							}
							
							// Set the value of the extended property to 0 (which is Sensitivity normal, 2 would be private)
							Message.ExtendedProperties[extendedPropertyindex].Value = 0;
							
							// Update the item on the server with the new client-side value of the target extended property.
							Message.Update(ConflictResolutionMode.AlwaysOverwrite);
						}
						extendedPropertyindex++;
					}					
				}
				catch(Exception ex) {
					log.Error("Error on update the item. Error message:", ex);
				}
			}	
		}
		
		public static ExchangeService ConnectToExchange(string MailboxID)
		{
			log.Info(string.Format("Connect to mailbox {0}",MailboxID));
			try {
				var service = new ExchangeService();
				service.UseDefaultCredentials = true;
				service.AutodiscoverUrl(MailboxID);
				service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, MailboxID);
				return service;
			}
			catch(Exception ex)	{
				log.Error("Connection to mailbox failed",ex);
			}
			return null;
		}
		
		public static string GetFolderPath(ExchangeService service, FolderId ID) {
			try {
				var FolderPathProperty = new ExtendedPropertyDefinition(0x66B5, MapiPropertyType.String);
				PropertySet psset1 = new PropertySet(BasePropertySet.FirstClassProperties);
				psset1.Add(FolderPathProperty);
				
				Folder FolderwithPath = Folder.Bind(service, ID, psset1);
				Object FolderPathVal = null;
				if (FolderwithPath.TryGetProperty(FolderPathProperty, out  FolderPathVal)) {
					// because the FolderPath contains characters we didn't want we need to fix that
					string FolderPathTemp = FolderPathVal.ToString();
					if (FolderPathTemp.Contains("￾")) {
						return FolderPathTemp.Replace("￾","\\");
					}
					else {
						return FolderPathTemp;
					}
				}
			} 
			catch (Exception ex) {
				log.Error("Failed to get folder path", ex);
			}
			return "";
		}
		
		public static FindFoldersResults Folders(ExchangeService service) {
			// try to find all folder that are unter MsgRootFolder
			int pageSize =  100;
			int pageOffset = 0;
			bool moreItems = true;
			var view = new FolderView(pageSize,pageOffset);
			view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
			view.PropertySet.Add(FolderSchema.DisplayName);
			// we define the seacht filter here. Find all folders which hold more than 0 elements
			SearchFilter searchFilter = new SearchFilter.IsGreaterThan(FolderSchema.TotalCount,0);
			view.Traversal = FolderTraversal.Deep;
			while (moreItems) {	
				try {
					findFolders = service.FindFolders(WellKnownFolderName.MsgFolderRoot, searchFilter,view);
					moreItems = findFolders.MoreAvailable;
					
					// if more folders than the offset is aviable we need to page
					if (moreItems) view.Offset += pageSize;
				}
				catch(Exception ex)	{
					log.Error("Failed to fetch folders.",ex);
					moreItems = false;
				}
			}
			return findFolders;
		}
		
		public static FindItemsResults<Item> PrivateItems(Folder Ordner) {
			int pageSize =  100;
			int pageOffset = 0;
			bool moreItems = true;
			var extendedPropertyDefinition = new ExtendedPropertyDefinition(0x36, MapiPropertyType.Integer);
			SearchFilter searchFilter = new SearchFilter.IsEqualTo(ItemSchema.Sensitivity, "Private");
			var view = new ItemView(pageSize, pageOffset);
			view.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.Sensitivity, ItemSchema.Subject, extendedPropertyDefinition);
			view.Traversal = ItemTraversal.Shallow;
			while (moreItems) {	
				try	{
					findResults = Ordner.FindItems(searchFilter, view);
					moreItems = findResults.MoreAvailable;
					
					// if more folders than the offset is aviable we need to page
					if (moreItems) view.Offset += pageSize;
				}
				catch(Exception ex)	{
					log.Error("Failed to fetch items.",ex);
					moreItems = false;
				}
			}
			return findResults;
		}	
	}
}