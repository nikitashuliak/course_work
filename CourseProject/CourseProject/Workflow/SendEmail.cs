using System;
using CourseProject.Log;
using CourseProject.Services;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace CourseProject.Workflow
{
    class SendEmail
    {
        public void SendLetter(Logger logger, string pathToFile, JToken inputs, string letterTemplate, string error, string originalFileName, List<Tuple<string, List<string>, int, int, List<Tuple<string, int>>, double, double>> data)
        {
            logger.Log(0, "Start sending message to Outlook. Letter Template is : " + letterTemplate + ".", false);
            try
            {
                if (letterTemplate == "1")
                {
                    sendLetterOfFirstTemplate(inputs, error);
                    logger.Log(0, "Successfully send letter of template 1.", false);
                }
                else if (letterTemplate == "2")
                {
                    sendLetterOfSecondTemplate(inputs, error, originalFileName);
                    logger.Log(0, "Successfully send letter of template 2.", false);
                }
                else if (letterTemplate == "3")
                {
                    sendLetterOfThirdTemplate(inputs, error, originalFileName);
                    logger.Log(0, "Successfully send letter of template 3.", false);
                }
                else if (letterTemplate == "4")
				{
                    sendLetterOfFourthTemplate(inputs, originalFileName, data);
                    logger.Log(0, "Successfully send letter of template 4.", false);
                } else
                {
                    logger.Log(1, "Unknown letter template. Email has not been send.", false);
                    return;
                }
            }
            catch (Exception ex)
            {
                logger.Log(1, ex.ToString() + " Execution stopped due to error.", false);
                return;
            }

        }

        private void sendLetterOfFirstTemplate(JToken inputs, string error)
        {
            string subject = "Execution stopped due to error";
            string message = @"Hello,<br></br><br></br>" +
                            "The process was not completed due to following error:<br></br>" +
                            error + "<br></br>" +
                            "Execution has been stopped.<br></br><br></br><br></br>" +
                            "Best,<br></br>" +
                            "Time Tracker Bot<br></br>" +
                            inputs["BotsName"].ToString() + "@deloitte.com<br></br>" +
                            "RPA";
            Outlook.SendMessage(inputs["EmailBusinessAdmins"].ToString(), String.Empty, String.Empty, subject, String.Empty,
                message, Convert.ToBoolean(inputs["SendMail"]), String.Empty);
        }

        private void sendLetterOfSecondTemplate(JToken inputs, string error, string originalFileName)
        {
            string subject = "Input file is invalid";
            string message = @"Hello,<br></br><br></br>" +
                            "The process was not completed due to invalid input file:<br></br>" +
                            error + "<br></br><br></br>" +
                            "Origanal File Name: " + originalFileName.Replace(".xlsx", "") + "<br></br><br></br><br></br>" +
                            "Best,<br></br>" +
                            "Time Tracker Bot<br></br>" +
                            inputs["BotsName"].ToString() + "@deloitte.com<br></br>" +
                            "RPA";

            Outlook.SendMessage(inputs["EmailBusinessAdmins"].ToString(), String.Empty, String.Empty, subject, String.Empty,
                message, Convert.ToBoolean(inputs["SendMail"]), String.Empty);
        }

        private void sendLetterOfThirdTemplate(JToken inputs, string error, string originalFileName)
		{
            string subject = "Error while creating statistick";
            string message = @"Hello, <br></br><br></br>" +
                              "The process has been stopped while creating statistick due to following error:<br></br>" +
                              error + "<br></br><br></br>" +
                              "Origanal File Name: " + originalFileName.Replace(".xlsx", "") + "<br></br><br></br><br></br>" +
                              "Best,<br></br>" +
                              "Time Tracker Bot<br></br>" +
                              inputs["BotsName"].ToString() + "@companyname.com<br></br>" +
                              "RPA";
            Outlook.SendMessage(inputs["EmailBusinessAdmins"].ToString(), String.Empty, String.Empty, subject, String.Empty,
                message, Convert.ToBoolean(inputs["SendMail"]), String.Empty);
        }

        private void sendLetterOfFourthTemplate(JToken inputs, string originalFileName, List<Tuple<string, List<string>, int, int, List<Tuple<string, int>>, double, double>> data)
        {
            string generatedTopStatistic = "";
            for (int i = 0; i < 3; i++)
			{
                generatedTopStatistic += CreateInfoByIndex(i, data);
			}
            string generatedNonTopStatistic = "";
            for (int i = data.Count-1; i > data.Count-4; i--)
            {
                generatedNonTopStatistic += CreateInfoByIndex(i, data);
            }
            string subject = "Statistick by TimeTracker Bot has been successfully created.";
            string message = @"Hello,<br></br><br></br>" +
                              "The validation has been completed according to provided input file - " + originalFileName.Replace(".xlsx", "") + ".<br></br><br></br>" +
                              "Here are the best 3 employees of previous month: <b>" + data[0].Item1 + "</b>, <b>" + data[1].Item1 + "</b>, <b>" + data[2].Item1 + "</b>.<br></br>" +
                              "<b>Statatistick</b>:<br></br><br></br>" +
                              generatedTopStatistic + "<br></br><br></br>" +
                              "Here are the worth 3 employees of previous month: <b>" + data[data.Count-1].Item1 + "</b>, <b>" + data[data.Count - 2].Item1 + "</b>, <b>" + data[data.Count - 3].Item1 + "</b>.<br></br>" +
                              "<b>Statatistick</b>:<br></br><br></br>" +
                              generatedNonTopStatistic + "<br></br><br></br>" +
                              "Best,<br></br>" +
                              "Time Tracker Bot<br></br>" +
                              inputs["BotsName"].ToString() + "@companyname.com<br></br>" +
                              "RPA";

            Outlook.SendMessage(inputs["EmailBusinessAdmins"].ToString(), String.Empty, String.Empty, subject, String.Empty,
                message, Convert.ToBoolean(inputs["SendMail"].ToString()), String.Empty);
        }

        private string CreateInfoByIndex(int i, List<Tuple<string, List<string>, int, int, List<Tuple<string, int>>, double, double>> data)
		{
            string text = "<b>" + data[i].Item1 + "</b>. Relevant information:<br></br>";
            string _relevantInfo = "";
            for (int t = 0; t < data[i].Item2.Count; t++)
			{
                if (t == 0)
				{
                    _relevantInfo += "<b>Position:</b> " + data[i].Item2[t] + "<br></br>";
				}
                if (t == 1)
				{
                    _relevantInfo += "<b>Specialization:</b> " + data[i].Item2[t] + "<br></br>";
                }
                if (t == 2)
				{
                    _relevantInfo += "<b>Team:</b> " + data[i].Item2[t] + "<br></br>";
                }
			}
            text += _relevantInfo;
            text += "Working time spent for previous month per project:<br></br>";
            string _timePerProjectInfo = "";
            for (int t = 0; t < data[i].Item5.Count; t++)
            {
                _timePerProjectInfo += data[i].Item5[t].Item1 + ": " + data[i].Item5[t].Item2.ToString() + "<br></br>";
            }
            text += _timePerProjectInfo;
            text += "<b>Total worked time:</b> " + data[i].Item3.ToString() + "<br></br>";
            text += "<b>Productivity coeficient:</b> " + String.Format("{0:0.0%}", data[i].Item6) + "<br></br><br></br>";
            return text;
        }
    }
}
