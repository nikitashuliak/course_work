using System;
using System.Collections.Generic;
using CourseProject.Log;

namespace CourseProject.Workflow
{
	class ReceivingStatisticsWorkflow
	{
		private static string err = "";
		private static string letterTemplate = "";
		private int workTimePerDay = 8;
		private List<string> nonWorkingProjectList = new List<string> { "Vacation/Day-off/Sick leave", "Company Name" };


		public static List<Tuple<string, List<string>, int, int, List<Tuple<string, int>>, double, double>> ProcessInputData(Logger globalLogger, List<Tuple<string, Tuple<List<string>, Dictionary<string, int>>>> data)
		{
			ReceivingStatisticsWorkflow receivingStatisticsWorkflow = new ReceivingStatisticsWorkflow();
			return receivingStatisticsWorkflow.GetStatisticsFromInputData(globalLogger, data);
		}

		private List<Tuple<string, List<string>, int, int, List<Tuple<string, int>>, double, double>> GetStatisticsFromInputData(Logger globalLogger, List<Tuple<string, Tuple<List<string>, Dictionary<string, int>>>> data)
		{
			globalLogger.Log(0, "Start processing input data.", false);

			List<Tuple<string, List<string>, int, int, List<Tuple<string, int>>, double, double>> result =
				new List<Tuple<string, List<string>, int, int, List<Tuple<string, int>>, double, double>>();

			try
			{
				var totalTime = GetTotalTime(data);
				if (totalTime % workTimePerDay != 0)
				{
					globalLogger.Log(1, $"Total time is not divisible by workTimePerDay value. Probably input data is not correct. Total time = {totalTime}, workTimePerDay = {workTimePerDay}. Computation continues.", false);
				}
				var totalWorkDaysInMonth = totalTime / workTimePerDay;

				Tuple<string, List<string>, int, int, List<Tuple<string, int>>, double, double> tmp;
				for (int i = 0; i < data.Count; i++)
				{
					// get name
					string _name = data[i].Item1;
					// get name info
					List<string> _nameInfo = data[i].Item2.Item1 as List<string>;
					int j = i;
					int totalWorkTime = 0;
					int totalNonWorkTime = 0;
					List<Tuple<string, int>> _timePerProjectInfo = new List<Tuple<string, int>>();
					while (j < data.Count && data[j].Item1 == _name)
					{
						// get total work time, total non work time, time per project info
						foreach (var el in data[j].Item2.Item2)
						{
							if (!nonWorkingProjectList.Contains(el.Key))
							{
								totalWorkTime += el.Value;
							}
							else
							{
								totalNonWorkTime += el.Value;
							}
							_timePerProjectInfo.Add(new Tuple<string, int>(el.Key, el.Value));
						}
						j++;
					}
					i = j - 1;
					result.Add(new Tuple<string, List<string>, int, int, List<Tuple<string, int>>, double, double>(
						_name,
						_nameInfo,
						totalWorkTime,
						totalNonWorkTime,
						_timePerProjectInfo,
						(double)totalWorkTime / totalTime,
						(double)totalNonWorkTime / totalTime
					));
				}

				result.Sort(delegate (Tuple<string, List<string>, int, int, List<Tuple<string, int>>, double, double> x1, Tuple<string, List<string>, int, int, List<Tuple<string, int>>, double, double> x2)
				{
					return x1.Item6.CompareTo(x2.Item6);
				});
				result.Reverse();
			}
			catch (Exception ex)
			{
				err = ex.Message;
				letterTemplate = "3";
				return null;
			}

			letterTemplate = "4";
			return result;
		}

		private int GetTotalTime(List<Tuple<string, Tuple<List<string>, Dictionary<string, int>>>> data)
		{
			int totalTime = 0;
			string _tempName = data[0].Item1;
			for (int i = 0; i < data.Count; i++)
			{
				if (data[i].Item1 == _tempName)
				{
					foreach (var el in data[i].Item2.Item2)
					{
						totalTime += el.Value;
					}
				} else
				{
					break;
				}
			}
			return totalTime;
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
