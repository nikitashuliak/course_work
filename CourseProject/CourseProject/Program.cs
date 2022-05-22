using System;
using CourseProject.Services;
using CourseProject.Log;
using CourseProject.Workflow;

namespace CourseProject
{
	class Program
	{
		private string err = "";
		private string letterTemplate = "";

		public void RunBot(string inputJsonFileName)
		{
			// try catch over initiation of logger
			var jsonData = FileManagment.GetInputs(inputJsonFileName);
			Logger globalLogger = new Logger(jsonData["GrantType"].ToString(), jsonData["Resource"].ToString(), jsonData["ClientId"].ToString(), jsonData["ClientSecret"].ToString(),
											jsonData["ServiceUrl"].ToString(), jsonData["LogUrl"].ToString(), jsonData["LogsFolder"].ToString(), jsonData["BotsName"].ToString(), Convert.ToBoolean(jsonData["DisableLogging"].ToString()));
			
			SendEmail Mail = new SendEmail();

			globalLogger.Log(0, "Execution has been started.", false);
			// get emails from outlook
			string[] outlookProcessingOutputs = InputsFromOutlook.OutlookProcess(globalLogger, jsonData);
			err = InputsFromOutlook.GetError();
			if (err.Trim() != "")
			{
				letterTemplate = InputsFromOutlook.GetLetterTemplate();
				Mail.SendLetter(globalLogger, outlookProcessingOutputs[1], jsonData, letterTemplate, err, outlookProcessingOutputs[0], null);
				return;
			}
			// validate input file
			var data = ValidationWorkflow.ValidationProcess(globalLogger, outlookProcessingOutputs[0], outlookProcessingOutputs[1]);
			if (data == null)
			{
				err = ValidationWorkflow.GetError();
				letterTemplate = ValidationWorkflow.GetLetterTemplate();
				Mail.SendLetter(globalLogger, outlookProcessingOutputs[1], jsonData, letterTemplate, err, outlookProcessingOutputs[0], null);
				return;
			}
			// process input data
			var resultData = ReceivingStatisticsWorkflow.ProcessInputData(globalLogger, data);
			if (resultData == null)
			{
				err = ReceivingStatisticsWorkflow.GetError();
				letterTemplate = ReceivingStatisticsWorkflow.GetLetterTemplate();
				Mail.SendLetter(globalLogger, outlookProcessingOutputs[1], jsonData, letterTemplate, err, outlookProcessingOutputs[0], null);
				return;
			}
			globalLogger.Log(0, "Input data has been processed successfully.", false);
			// send result to recipients
			err = ReceivingStatisticsWorkflow.GetError();
			letterTemplate = ReceivingStatisticsWorkflow.GetLetterTemplate();
			Mail.SendLetter(globalLogger, outlookProcessingOutputs[1], jsonData, letterTemplate, err, outlookProcessingOutputs[0], resultData);
			globalLogger.Log(0, "Execution has been completed.", false);
		}
	}
}
