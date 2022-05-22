using System;
using System.Linq;
using System.Data;
using System.Collections.Generic;
using System.Management;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace CourseProject.Services
{
	class Excel
	{
		public static StringComparer comparer = StringComparer.InvariantCultureIgnoreCase;
		public static Dictionary<string, Application> instances = new Dictionary<string, Application>();
		public static Dictionary<string, Workbook> workbooks = new Dictionary<string, Workbook>(comparer);
		public static Dictionary<string, Dictionary<string, Worksheet>> wbWorksheets = new Dictionary<string, Dictionary<string, Worksheet>>();

		public static Application ExcelApp = null;
		public static Workbook wbBrd = null;
		public static Worksheet wsBrd = null;

		public static void CloseAllExcelInstances()
		{
			try
			{
				KillProcess("EXCEL.EXE");
			}
			catch (Exception ex)
			{
				throw new SystemException(ex.ToString());
			}


			wbWorksheets.Clear();
			workbooks.Clear();
			instances.Clear();
		}

		public static void KillProcess(string ProcessName)
		{
			//get current user
			string ProcessUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
			//get all process by name
			Process[] foundProcesses = Process.GetProcessesByName(ProcessName);

			string strMessage = string.Empty;
			foreach (Process p in foundProcesses)
			{
				//get process owner
				string UserName = GetProcessOwner(p.Id);

				strMessage = string.Format("Process Name: {0} | Process ID: {1} | User Name : {2} | StartTime {3}",
												 p.ProcessName, p.Id.ToString(), UserName, p.StartTime.ToString());
				//compare process user with current system user
				bool PrcoessUserName_Is_Matched = UserName.Equals(ProcessUserName);

				if ((ProcessUserName.ToLower() == "all") ||
					 PrcoessUserName_Is_Matched)
				{
					p.Kill();

				}
			}
		}

		public static string GetProcessOwner(int processId)
		{
			string query = "Select * From Win32_Process Where ProcessID = " + processId;
			ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
			ManagementObjectCollection processList = searcher.Get();

			foreach (ManagementObject obj in processList)
			{
				string[] argList = new string[] { string.Empty, string.Empty };
				int returnVal = Convert.ToInt32(obj.InvokeMethod("GetOwner", argList));
				if (returnVal == 0)
				{
					return argList[1] + "\\" + argList[0];   // return DOMAIN\user
				}
			}
			return "NO OWNER";
		}

		public static void KillExcelProcess()
		{
			System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
			foreach (System.Diagnostics.Process p in process)
			{
				if (!string.IsNullOrEmpty(p.ProcessName))
				{
					try
					{
						//Console.WriteLine();
						p.Kill();
					}
					catch { }
				}
			}
		}

		public static int findRowWithHeaders(List<string> requiredHeadersList, ref object[,] data)
		{
			int rowIndex = -1;

			for (int index1 = 1; index1 <= data.GetLength(0); ++index1)
			{
				for (int index2 = 1; index2 <= data.GetLength(1); ++index2)
				{
					try
					{
						if (data[index1, index2] == null) continue;
						if (requiredHeadersList.Contains(data[index1, index2].ToString().Trim().ToLower()))
						{
							if (rowIndex < 0)
							{
								rowIndex = index1;
								break;
							}
						}
					}
					catch {
						Console.WriteLine("Some error");
					}
				}
				if (rowIndex >= 0)
				{
					break;
				}
			}

			return rowIndex;
		}

		public static int countRequiredHeaders(Microsoft.Office.Interop.Excel.Worksheet ws, List<string> requiredHeadersList)
		{
			int reqHeadersCount = 0;
			var nInLastCol = 0;
			var nInLastRow = 0;
			try
			{
				nInLastCol = ws.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value,
					System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
					false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
				nInLastRow = 1000;
			}
			catch
			{
			}

			if (nInLastCol > 0 && nInLastRow > 0)
			{
				Microsoft.Office.Interop.Excel.Range range = ws.get_Range("A1", getExcelColumnName(nInLastCol) + nInLastRow);

				int rangeRowsCount = range.Rows.Count;
				int rangeColumnsCount = range.Columns.Count;
				int rowIndex = 0;
				object[,] data = (object[,])range.Value2;
				for (int index1 = 1; index1 <= rangeRowsCount; ++index1)
				{
					for (int index2 = 1; index2 <= rangeColumnsCount; ++index2)
					{
						try
						{
							if (requiredHeadersList.Contains(System.Convert.ToString(data[index1, index2]).Trim().ToLower()))
							{
								//++reqHeadersCount;
								if (rowIndex == 0)
								{
									rowIndex = index1;
									break;
								}
							}
						}
						catch { }
					}
					if (rowIndex > 0)
						break;
				}
				if (rowIndex > 0)
				{
					for (int i = 1; i < rangeColumnsCount; ++i)
					{
						if (requiredHeadersList.Contains(System.Convert.ToString(data[rowIndex, i]).Trim().ToLower()))
						{
							++reqHeadersCount;
						}
					}
				}
			}

			return reqHeadersCount;
		}

		public static object[,] readWorksheetToArray(Microsoft.Office.Interop.Excel.Worksheet ws)
		{
			var nInLastCol = 0;
			var nInLastRow = 0;
			object[,] data = null;
			try
			{
				nInLastCol = ws.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value,
					System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
					false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
				nInLastRow = ws.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value,
					System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious,
					false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
			}
			catch
			{
			}
			if (nInLastCol > 0 & nInLastRow > 0)
				data = ws.get_Range("A1", getExcelColumnName(nInLastCol) + nInLastRow).Value2 as object[,];

			return data;
		}

		public static string getExcelColumnName(int columnNumber)
		{
			int dividend = columnNumber;
			string columnName = String.Empty;
			int modulo;

			while (dividend > 0)
			{
				modulo = (dividend - 1) % 26;
				columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
				dividend = (int)((dividend - modulo) / 26);
			}
			return columnName;
		}

		public static List<Tuple<string, Tuple<List<string>, Dictionary<string, int>>>> GetValuesUnderHeaders(
			string wbName, string sheetName, List<string> requiredHeadrsList)
		{
			// result is saved in object below
			List<Tuple<string, Tuple<List<string>, Dictionary<string, int>>>> data = new List<Tuple<string, Tuple<List<string>, Dictionary<string, int>>>>();

			Worksheet sheet = (Worksheet)ExcelApp.Workbooks[wbName].Sheets[sheetName];
			Range range = sheet.UsedRange;
			object[,] objectArray = (object[,])range.Value[XlRangeValueDataType.xlRangeValueDefault];
			int rows = range.Rows.Count;
			int cols = range.Columns.Count;
			int lastWorkDayInMonth = 0;
			int start_col = 0, end_col = 0;
			bool headrsFoundFlag = false;
			bool foundLastRow = false;

			for (int r = 1; r <= rows; r++)
			{
				if (foundLastRow) break;
				if (!headrsFoundFlag)
				{
					for (int c = 1; c <= cols; c++)
					{
						if (objectArray[r, c] == null) continue;
						if (requiredHeadrsList.Contains(objectArray[r, c].ToString().Trim().ToLower()))
						{
							start_col = c;
							headrsFoundFlag = true;
							int cur = c;
							while (objectArray[r, cur] != null)
							{
								cur++;
							}
							lastWorkDayInMonth = cur - 1;
							end_col = cur;
							break;
						}
					}
				}
				else
				{
					if (objectArray[r, start_col] == null)
					{
						foundLastRow = true;
						break;
					}
					string _name = "";
					List<string> _info = new List<string>();
					Dictionary<string, int> workInfo = new Dictionary<string, int>();
					int mid_col = start_col + 4;
					for (int c = start_col; c < mid_col; c++)
					{
						if (c == start_col) _name = objectArray[r, c].ToString();
						else _info.Add(objectArray[r, c].ToString());
					}
					string workName = objectArray[r, mid_col].ToString();
					workInfo.Add(workName, 0);
					for (int c = mid_col + 1; c < end_col; c++)
					{
						if (objectArray[r, c] == null) continue;
						else workInfo[workName] += Int32.Parse(objectArray[r, c].ToString());
					}
					data.Add(new Tuple<string, Tuple<List<string>, Dictionary<string, int>>>(_name, new Tuple<List<string>, Dictionary<string, int>>(_info, workInfo)));
				}
			}

			return data;
		}

		public static void CloseWorkbook(string wbName, bool saveChanges)
		{
			if (ExcelApp.Workbooks[wbName] == null)
				throw new ArgumentException(string.Format("Workbook '{0}' does not exist.", wbName));
			if (saveChanges)
				ExcelApp.Workbooks[wbName].Save();
			ExcelApp.Workbooks[wbName].Close();
		}

		public static string CreateWorkbook(string pathToLocalFolder, string fileName)
		{
			Workbook newWB = ExcelApp.Workbooks.Add();
			newWB.SaveAs(Path.Combine(pathToLocalFolder, fileName));
			return newWB.FullName;
		}

		public static void CreateInstance()
		{
			ExcelApp = new Application();
			ExcelApp.Visible = false;
			ExcelApp.DisplayAlerts = false;
		}

		public static string OpenWorkbook(string wbPath)
		{
			string wbName = "";
			Microsoft.Office.Interop.Excel.Workbook wb = null;
			if (!File.Exists(wbPath))
				throw new ArgumentException(string.Format("File '{0}' does not exist.", wbPath));
			try
			{
				wb = Excel.ExcelApp.Workbooks.Open(Filename: wbPath, ReadOnly: false, Password: "", WriteResPassword: "", Origin: XlPlatform.xlWindows);
				wbName = wb.Name;
			}
			catch (COMException ex)
			{
				if (ex.ToString().Contains("file format or file extension is not valid"))
					throw new ArgumentException(string.Format("{0} is not valid due to {1}", Path.GetFileName(wbPath), "the file is broken (cannot be read)"));
				if (ex.ToString().Contains("The password"))
					throw new ArgumentException(string.Format("{0} is not valid due to {1}", Path.GetFileName(wbPath), "the file is password protected"));
			}

			return wbName;
		}

		public static void WriteData(System.Data.DataTable data, string wbName, string wsName)
		{
			object[,] dataToWrite = new object[data.Rows.Count + 1, data.Columns.Count];
			for (int i = 0; i < data.Columns.Count; ++i)
				dataToWrite[0, i] = data.Columns[i].ColumnName;

			for (int i = 0; i < data.Rows.Count; ++i)
			{
				for (int j = 0; j < data.Columns.Count; ++j)
					dataToWrite[i + 1, j] = data.Rows[i][j];
			}

			Workbook wb = ExcelApp.Workbooks[wbName];
			Worksheet ws = (Worksheet)ExcelApp.Workbooks[wbName].Sheets[wsName];
			ws.Range["A1", getExcelColumnName(dataToWrite.GetLength(1)) + dataToWrite.GetLength(0)].Value2 = dataToWrite;

			ws.Columns.AutoFit();


		}

		public static string GetSheetName(string wbName, List<string> requiredHeadersList)
		{
			string error = "";
			int rowIndex = -1;
			string sheetName = "";
			System.Data.DataTable dt = new System.Data.DataTable();

			try
			{
				Dictionary<string, int> worksheetsHeaders = new Dictionary<string, int>();
				for (int i = 1; i <= (int)Excel.ExcelApp.Workbooks[wbName].Sheets.Count; ++i)
				{
					worksheetsHeaders.Add(((Worksheet)(Excel.ExcelApp.Workbooks[wbName].Sheets[i])).Name,
						Excel.countRequiredHeaders((Worksheet)Excel.ExcelApp.Workbooks[wbName].Sheets[i], requiredHeadersList));
				}

				string wsName = worksheetsHeaders.Aggregate((l, r) => l.Value > r.Value ? l : r).Key;
				DataRow dtRow = dt.NewRow();

				Worksheet xlWorkSheet = (Worksheet)Excel.ExcelApp.Workbooks[wbName].Sheets[wsName];
				object[,] elements = Excel.readWorksheetToArray(xlWorkSheet);

				if (elements != null && (elements.GetLength(0) > 0 & elements.GetLength(1) > 0))
				{

					rowIndex = Excel.findRowWithHeaders(requiredHeadersList, ref elements);
					Dictionary<string, int> headersCount = new Dictionary<string, int>();
					Dictionary<string, bool> emptyColumns = new Dictionary<string, bool>();
					if (rowIndex >= 0)
					{
						sheetName = xlWorkSheet.Name;
						dt.Rows.Add(dtRow);
					}
				}
			}
			catch (System.Exception ex)
			{
				Console.WriteLine(ex.Message);
				error = ex.ToString();
			}

			return sheetName;
		}
	}
}
