using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using System.IO;
using Newtonsoft.Json;


namespace CourseProject.Services
{
	struct Pair
	{
		private string str;
		private bool flag;

		public void SetFirst(string text)
		{
			str = text;
		}

		public void SetSecond(bool value)
		{
			flag = value;
		}

		public string GetFirst()
		{
			return str;
		}

		public bool GetSecond()
		{
			return flag;
		}
	}

	class FileManagment
	{
		public static HashSet<string> globalSet = new HashSet<string>();

		public static JToken GetInputs(string inputFilePath)
		{
			StreamReader r = new StreamReader(inputFilePath);
			string jsonString = r.ReadToEnd();
			JObject data = JsonConvert.DeserializeObject<JObject>(jsonString);
			//JObject data = JObject.Parse(File.ReadAllText(inputFilePath));
			return data["InputData"];
		}
		public static Pair CheckDownloading(string fileName)
		{
			string downloadFolder = @"C:\Users\{0}\Downloads";
			downloadFolder = String.Format(downloadFolder, Environment.UserName);
			fileName = fileName.Substring(0, fileName.Length - 5);
			//Console.WriteLine(fileName);
			Pair pair = new Pair();
			pair.SetFirst(""); pair.SetSecond(false);
			var files = Directory.GetFiles(downloadFolder, fileName + "*.xlsx");
			//Console.WriteLine(files.Length);
			foreach (var tFile in files)
			{
				//Console.WriteLine(tFile);
				if (!globalSet.Contains(tFile))
				{
					pair.SetFirst(tFile);
					pair.SetSecond(true);
				}
			}
			//Console.WriteLine(pair.GetFirst() + " " + pair.GetSecond());
			return pair;
		}

		public static Pair MoveFile(string path, string toPath, string newName)
		{
			path = path.Trim();
			toPath = toPath.Trim('\\').Trim();
			Pair pair = new Pair();
			pair.SetFirst(""); pair.SetSecond(false);
			try
			{
				if (!File.Exists(path))
					throw new ApplicationException("The specified file at " + path + " does not exist.");
				if (!Directory.Exists(toPath))
					throw new ApplicationException("The specified directory at " + path + " does not exist.");
				FileInfo file = new FileInfo(path);
				string name = file.Name;
				if (String.IsNullOrEmpty(newName))
				{
					File.Move(path, Path.Combine(toPath, file.Name));
				}
				else
				{
					File.Move(path, Path.Combine(toPath, newName));
					name = newName;
				}
				if (File.Exists(Path.Combine(toPath, name)))
				{
					pair.SetFirst(Path.Combine(toPath, name));
					pair.SetSecond(true);
				}
			}
			catch (Exception ex)
			{
				pair.SetFirst(ex.ToString());
				pair.SetSecond(false);
			}

			return pair;
		}

	}
}
