using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using Newtonsoft.Json.Linq;
using System.IO;


namespace CourseProject.Log
{
	public class Logger
	{
		public static string ServiceUrl;
		public static Uri LogUrl;
		public static Dictionary<string, string> ReqTokenParams;
		public static string LogsFolder;
		public static string BotName;
		public static bool DisableLogging;
		public enum LogLevel
		{
			Information = 0,
			Error = 1
		}
		private static HttpClient httpClient = new HttpClient();
		private static Dictionary<string, string> tokenData = new Dictionary<string, string>
		{
			{ "token_type", "" },
			{ "access_token", "" },
			{ "expires_on", "-1" }
		};

		/*
		 * Sends a log message to RPA Commander
		 */
		private void logToRpaCommander(LogLevel logLevel, string message)
		{
			if (isTokenExpired())
				getTokenData();

			var reqBody = new StringContent(getJsonLogStr(logLevel, message), Encoding.UTF8, "application/json");

			var reqMessage = new HttpRequestMessage
			{
				RequestUri = LogUrl,
				Method = HttpMethod.Post,
				Content = reqBody
			};

			reqMessage.Headers.Authorization = new AuthenticationHeaderValue(tokenData["token_type"], tokenData["access_token"]);

			HttpResponseMessage responseMessage = null;
			try
			{
				responseMessage = httpClient.SendAsync(reqMessage).Result;
				responseMessage.EnsureSuccessStatusCode();

				if (responseMessage.Content.Headers.ContentLength != 0)
					throw new HttpRequestException(string.Format("Log Message '{0}' was not sent to 'RPA Commander'", message));
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.ToString());
				logToFile((LogLevel)logLevel, ex.ToString());
			}
			finally
			{
				if (responseMessage != null)
					responseMessage.Dispose();
			}
		}

		/*
		 * Writes a log message to log file
		 */
		private void logToFile(LogLevel logLevel, string message)
		{
			string logFileName = string.Join("_", Environment.UserName, DateTime.UtcNow.ToString("MMddyyyy") + ".log");
			string logFilePath = Path.Combine(LogsFolder, "logs", BotName, logFileName);

			string formatedLogMessage = string.Format("{0}\t[{1}]\t{2}", DateTime.UtcNow, logLevel, message);

			try
			{
				if (!Directory.Exists(Path.GetDirectoryName(logFilePath)))
					Directory.CreateDirectory(Path.GetDirectoryName(logFilePath));

				File.AppendAllText(logFilePath, formatedLogMessage.ToString() + Environment.NewLine);
			}
			catch (Exception ex)
			{
				//Send exception text to VSTS
				Console.WriteLine(ex.ToString());
			}
		}

		/*
		 * Gets an auth token
		 */
		private void getTokenData()
		{
			HttpResponseMessage responseMessage = null;
			try
			{
				responseMessage = httpClient.PostAsync(ServiceUrl, new FormUrlEncodedContent(ReqTokenParams)).Result;
				responseMessage.EnsureSuccessStatusCode();

				JObject responseJson = JObject.Parse(responseMessage.Content.ReadAsStringAsync().Result);

				tokenData["token_type"] = responseJson.SelectToken("token_type").Value<string>();
				tokenData["access_token"] = responseJson.SelectToken("access_token").Value<string>();
				tokenData["expires_on"] = responseJson.SelectToken("expires_on").Value<string>();
			}
			catch (Exception ex)
			{
				Console.WriteLine("Unable to get an auth token for RPA Commander");
				Console.WriteLine(ex.ToString());
			}
			finally
			{
				if (responseMessage != null)
					responseMessage.Dispose();
			}
		}

		/*
		 * Check if token has expired
		 */
		private bool isTokenExpired()
		{
			bool expired = false;

			if (DateTimeOffset.Compare(DateTimeOffset.FromUnixTimeMilliseconds(Convert.ToInt64(tokenData["expires_on"])), DateTimeOffset.UtcNow.AddMinutes(-1)) < 0)
				expired = true;

			return expired;
		}

		/*
		 * Returns JSON string that will be sent to RPA Commander as part of request content
		 */
		private string getJsonLogStr(LogLevel logLevel, string message)
		{
			var reqJson = new JObject(
				new JProperty("Id", Guid.NewGuid()),
				new JProperty("Time", DateTime.UtcNow.ToString("MM.dd.yyyy HH:mm:ss")),
				new JProperty("Level", logLevel),
				new JProperty("Data", message)
			);

			return reqJson.ToString();
		}


		public void Close()
		{
			if (httpClient != null)
				httpClient.Dispose();
		}
		public Logger(string grantType, string resource, string clientId, string clientSecret, string serviceUrl, string logUrl, string logsFolder, string botName, bool disableLogging)
		{
			try
			{
				ReqTokenParams = new Dictionary<string, string>
				{
					{ "grant_type", grantType },
					{ "resource", resource },
					{ "client_id", clientId },
					{ "client_secret", clientSecret }
				};
				ServiceUrl = serviceUrl;
				LogUrl = new Uri(logUrl);
				LogsFolder = logsFolder;
				BotName = botName;
				DisableLogging = disableLogging;
			}
			catch (Exception ex)
			{
				throw new System.Exception(ex.ToString());
			}

		}

		public void Log(int logLevel, string message, bool sendToRpaCommander)
		{
			if (!DisableLogging)
			{
				Console.WriteLine("{0}\t[{1}]\t{2}", DateTime.UtcNow, (LogLevel)logLevel, message);
				try
				{
					logToFile((LogLevel)logLevel, message);
				}
				catch
				{
					Console.WriteLine("{0}\t[{1}]\t Unable to write log message to the file...", DateTime.UtcNow, (LogLevel)logLevel);
				}
				if (sendToRpaCommander)
					try
					{
						logToRpaCommander((LogLevel)logLevel, message);
					}
					catch
					{
						Console.WriteLine("{0}\t[{1}]\t Unable to send message to the RPA Commander...", DateTime.UtcNow, (LogLevel)logLevel);
					}
			}
		}

	}
}
