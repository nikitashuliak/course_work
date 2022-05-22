using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Office.Interop.Outlook;

namespace CourseProject.Services
{
	class Outlook
	{
		public static Application outlookApp = new Application();
		public static NameSpace outlookNS = outlookApp.GetNamespace("MAPI");
		public static MAPIFolder oFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
		public static Items oItems = oFolder.Items;
		public static MailItem newItem = null;

		public static DataTable GetMessagesFromOutlook(bool includeRead, bool includeUnread, string senderEmails, string subject, string subfolder, string receivedLatest, string receivedEarliest)
		{
			outlookApp = new Application();
			outlookNS = outlookApp.GetNamespace("MAPI");
			oFolder = (MAPIFolder)outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Parent;
			Items messages = null;
			if (subfolder != "")
			{
				foreach (string name in subfolder.Split('\\'))
				{
					oFolder = oFolder.Folders[name];
				}
			}
			else
				oFolder = outlookNS.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

			oItems = oFolder.Items;
			StringBuilder builder = new StringBuilder();
			builder.Append("@SQL=");
			if (includeRead & !includeUnread)
				builder.Append("urn:schemas:httpmail:read=1 AND ");
			if (includeUnread & !includeRead)
				builder.Append("urn:schemas:httpmail:read=0 AND ");
			if (subject.Trim() != "")
				builder.Append(string.Format("urn:schemas:httpmail:subject LIKE '%{0}%' AND ", subject.Trim()));
			if (receivedEarliest.Trim() != "")
				builder.Append(string.Format("urn:schemas:httpmail:datereceived>='{0}' AND ", Convert.ToDateTime(receivedEarliest)));
			if (receivedLatest.Trim() != "")
				builder.Append(string.Format("urn:schemas:httpmail:datereceived<='{0}' AND ", Convert.ToDateTime(receivedLatest)));
			if (builder.ToString() != "@SQL=")
			{
				string filterCondition = builder.ToString().TrimEnd(new char[] { ' ', 'A', 'N', 'D' });
				messages = oItems.Restrict(filterCondition);
			}
			else
				messages = oItems;
			DataTable tmpItems = convertOutlookItemsToDataTable(messages);
			DataTable items = tmpItems.Clone();
			if (senderEmails.Trim() != "")
			{
				string[] senderEmailsArr = senderEmails.Split(';');
				int index = tmpItems.Columns.IndexOf("SenderEmailAddress");
				foreach (string name in senderEmailsArr)
				{
					for (int i = 0; i < tmpItems.Rows.Count; ++i)
					{
						if (tmpItems.Rows[i][index] != DBNull.Value && ((string)tmpItems.Rows[i][index]).Trim().ToLower() == name.Trim().ToLower())
						{
							DataRow r = items.NewRow();
							r.ItemArray = tmpItems.Rows[i].ItemArray;
							items.Rows.Add(r);
						}
					}
				}
			}
			else
				items = tmpItems.Copy();

			return items;
		}

		public static DataTable convertOutlookItemsToDataTable(Items items)
		{
			DataTable dataTable = new DataTable();
			dataTable.Columns.Add("ItemID");
			dataTable.Columns.Add("To");
			dataTable.Columns.Add("SenderEmailAddress");
			dataTable.Columns.Add("Subject");
			dataTable.Columns.Add("Message");
			dataTable.Columns.Add("Unread");
			dataTable.Columns.Add("ReceivedTime");
			dataTable.Columns.Add("Attachments");
			foreach (Object it in items)
			{
				if (it is MailItem)
				{
					List<object> list = new List<object>();
					list.Add(((MailItem)it).EntryID);
					list.Add(((MailItem)it).To);
					if (((MailItem)it).SenderEmailType == "EX")
					{
						list.Add(convertExEmailToSMPT((MailItem)it));
					}
					else
						list.Add(((MailItem)it).SenderEmailAddress);

					list.Add(((MailItem)it).Subject);
					list.Add(((MailItem)it).Body);
					list.Add(((MailItem)it).UnRead);
					list.Add(((MailItem)it).ReceivedTime);
					if (((MailItem)it).Attachments.Count > 0)
					{
						List<string> attNames = new List<string>();
						foreach (Attachment a in ((MailItem)it).Attachments)
						{
							if (a.Type == OlAttachmentType.olByValue)
								attNames.Add(a.FileName);
						}
						if (attNames.Count > 0)
						{
							string names = attNames.Aggregate((x, y) => x + "|" + y);
							list.Add(names);
						}
						else
							list.Add("");
					}
					else
						list.Add("");
					dataTable.Rows.Add(list.ToArray());
				}
			}
			return dataTable;
		}

		public static string convertExEmailToSMPT(MailItem olMail)
		{
			string senderEmail = "";
			var objReply = ((_MailItem)olMail).Reply();
			Recipient objRecipient = objReply.Recipients[1];
			string strEntryId = objRecipient.EntryID;
			AddressEntry objAddressentry = outlookNS.GetAddressEntryFromID(strEntryId);
			if (objAddressentry != null)
			{
				ExchangeUser eu = objAddressentry.GetExchangeUser();
				if (eu != null)
					senderEmail = eu.PrimarySmtpAddress;
			}
			return senderEmail;
		}

		public static string GetOriginalAttachmentName(DataRow row)
		{
			return row.ItemArray[row.ItemArray.Length - 1].ToString();
		}

		public static string DonwloadAttachment(string itemId, string fileName)
		{
			string pathToLocalFolder = @"C:\Workfolder\SavedAttachments";

			int prevCount = GetFilesBeforeDownloading(fileName);
			string pathToDownload = CreatePathToDownloadingFolder();
			SaveAttachment(fileName, pathToDownload, itemId);
			Pair pair = FileManagment.CheckDownloading(fileName);
			if (!pair.GetSecond())
			{
				return String.Empty;
			}
			pair = FileManagment.MoveFile(pair.GetFirst(), pathToLocalFolder, fileName);
			return pair.GetFirst();
		}

		public static int GetFilesBeforeDownloading(string fileName)
		{
			string downloadFolder = @"C:\Users\{0}\Downloads";
			downloadFolder = String.Format(downloadFolder, Environment.UserName);
			int count = 0;
			try
			{
				var files = Directory.GetFiles(downloadFolder, fileName + "*.xlsx");
				foreach (var tFile in files)
				{
					FileManagment.globalSet.Add(tFile);
				}
				count = FileManagment.globalSet.Count;
			}
			catch
			{
				count = -1;
			}

			return count;
		}

		public static string CreatePathToDownloadingFolder()
		{
			string downloadFolder = @"C:\Users\{0}\Downloads";
			string path = String.Format(downloadFolder, Environment.UserName);
			return path;
		}

		public static void SaveAttachment(string fileName, string pathToFolder, string itemId)
		{
			outlookApp = new Application();
			outlookNS = outlookApp.GetNamespace("MAPI");
			MailItem item = (MailItem)outlookNS.GetItemFromID(itemId);
			Attachments attachments = item.Attachments;
			List<string> fileNames = null;
			if (fileName != "*")
			{
				fileNames = fileName.Split(';').ToList();
			}
			foreach (Attachment attach in attachments)
			{
				if (fileName != "*")
				{
					foreach (string val in fileNames)
					{
						if (attach.FileName.Contains(val))
							attach.SaveAsFile(Path.Combine(pathToFolder, attach.FileName));
					}
				}
				else
					attach.SaveAsFile(Path.Combine(pathToFolder, attach.FileName));
			}
		}

		public static void SendMessage(string to, string bcc, string cc, string subject, string attachments, string message, bool send, string sendOnBehalf)
		{
			outlookApp = new Application();
			outlookNS = outlookApp.GetNamespace("MAPI");
			MailItem item = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);
			SetMailAttributes(to, bcc, cc, subject, attachments, message, sendOnBehalf, ref item);
			if (send)
				((_MailItem)item).Send();
		}

		public static void SetMailAttributes(string to, string bcc, string cc, string subject, string attachments, string message, string sendOnBehalf, ref MailItem item)
		{
			if (to.Trim() == "")
				throw new System.Exception("Recipient email address is not mentioned!");
			else
			{
				item.Subject = subject;
				SetRecipients(to, bcc, cc, item);
				if (attachments.Trim() != "")
				{
					Attachments mailAttachments = item.Attachments;
					List<string> attNames = attachments.Split(';').ToList();
					//remove empty paths from attachments
					for (int i = 0; i < attNames.Count; ++i)
					{
						if (attNames[i].Trim() == "")
						{
							attNames.RemoveAt(i);
							--i;
						}
					}
					foreach (string name in attNames)
					{
						mailAttachments.Add(name);
					}
				}
				if (sendOnBehalf != "")
					item.SentOnBehalfOfName = sendOnBehalf;
				if (message.Trim() != "")
					item.HTMLBody = message;
				item.Save();
			}
		}

		public static void SetRecipients(string to, string bcc, string cc, MailItem item)
		{
			Dictionary<string, string> outAddresses = new Dictionary<string, string>();
			outAddresses.Add("to", to);
			if (bcc != "")
				outAddresses.Add("bcc", bcc);
			if (cc != "")
				outAddresses.Add("cc", cc);
			for (int i = 0; i < outAddresses.Count; ++i)
			{
				List<string> addresses = outAddresses.Values.ElementAt(i).Split(';').ToList();
				if (addresses.Count > 0)
				{
					foreach (string val in addresses)
					{
						Recipient rec = item.Recipients.Add(val);
						rec.Resolve();
						if (rec.Resolved)
						{
							if (outAddresses.Keys.ElementAt(i) == "to")
								rec.Type = 1;
							if (outAddresses.Keys.ElementAt(i) == "bcc")
								rec.Type = 3;
							if (outAddresses.Keys.ElementAt(i) == "cc")
								rec.Type = 2;
						}
					}
				}
				addresses.Clear();
			}
		}

		public static void MarkMail(string itemID, bool markAsUnread, bool markAsRead)
		{
			outlookApp = new Application();
			outlookNS = outlookApp.GetNamespace("MAPI");
			MailItem item = (MailItem)outlookNS.GetItemFromID(itemID);
			if (markAsUnread)
				item.UnRead = true;
			if (markAsRead)
				item.UnRead = false;
		}
	}
}
