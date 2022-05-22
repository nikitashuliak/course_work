using System;
using System.Data;
using Newtonsoft.Json.Linq;
using CourseProject.Log;
using CourseProject.Services;

namespace CourseProject.Workflow
{

	class InputsFromOutlook
	{
		private const string Reasons = @"Attachment is invalid for some of the reasons: 1) There is no attachment in email; 2) There are more than one attachmetns; 3) Attachment format is not appropriate.";
		private const bool MarkAsUnread = true, MarkAsRead = false;
		private static string err = "";
		private static string letterTemplate = "";
		public string[] GetInputsFromOutlook(Logger globalLogger, JToken data)
		{
			globalLogger.Log(0, "Search for new messages.", false);
			DataTable messagesFromOutlook = new DataTable();
			int cntMessages = 0;
			try
			{
				messagesFromOutlook = Outlook.GetMessagesFromOutlook(false, true, String.Empty, data["EmailSubject"].ToString(), String.Empty, String.Empty, String.Empty);
				cntMessages = messagesFromOutlook.Rows.Count;
			}
			catch (Exception ex)
			{
				//error = ex.ToString();
				//dt["letterTemplate"] = "1";
				globalLogger.Log(1, ex.ToString(), false);
				throw new Exception(ex.ToString());
			}

			if (cntMessages == 0)
			{
				/*dt["error"] = "None messages received.";
				dt["letterTemplate"] = "1";*/
				err = "None messages received.";
				letterTemplate = "1";
				globalLogger.Log(1, err, false);
				//throw new SystemException(error);
				return new string[] { "" };
			}
			DataRow earliestMessage = messagesFromOutlook.Rows[cntMessages - 1];
			Outlook.MarkMail(earliestMessage.ItemArray[0].ToString(), MarkAsUnread, MarkAsRead);

			if (earliestMessage.ItemArray[earliestMessage.ItemArray.Length - 1].ToString() == ""
				|| earliestMessage.ItemArray[earliestMessage.ItemArray.Length - 1].ToString().Contains("|")
				|| !earliestMessage.ItemArray[earliestMessage.ItemArray.Length - 1].ToString().Contains(".xlsx"))
			{
				err = Reasons;
				letterTemplate = "1";
				globalLogger.Log(1, err, false);
				//throw new SystemException(reasons);
				return new string[] { "" };
			}

			string originalFileName = Outlook.GetOriginalAttachmentName(earliestMessage);

			string pathToRetrievedFile = Outlook.DonwloadAttachment(earliestMessage.ItemArray[0].ToString(),
													earliestMessage.ItemArray[earliestMessage.ItemArray.Length - 1].ToString());



			if (pathToRetrievedFile.Trim() == "")
			{
				err = @"Error while downloading file";
				letterTemplate = "1";
				globalLogger.Log(1, err, false);
				//throw new SystemException(error);
				return new string[] { "" };
			}
			// MARK As READ

			globalLogger.Log(0, "Input File retrieved.", false);

			/*dt["originalFileName"] = originalFileName;
			dt["pathToRetrievedFile"] = pathToRetrievedFile;*/
			return new string[] { originalFileName, pathToRetrievedFile };
		}

		public static string[] OutlookProcess(Logger globalLogger, JToken data)
		{
			InputsFromOutlook inputs = new InputsFromOutlook();
			return inputs.GetInputsFromOutlook(globalLogger, data);
		}

		public static string GetError()
		{
			return err;
		}

		public static string GetLetterTemplate()
		{
			return letterTemplate;
		}
	}
}
