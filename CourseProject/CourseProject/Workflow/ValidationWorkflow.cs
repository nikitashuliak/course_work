using System;
using System.Collections.Generic;
using CourseProject.Services;
using CourseProject.Log;

namespace CourseProject.Workflow
{
    class ValidationWorkflow
    {
        private const string ListOfReasons = @"- file is broken
                                        - file does not contain data after header (on all tabs)
                                        - file does not contain header (on all tabs)
                                        - all tabs are locked
                                        - file is password protected";
        private static string err = "";
        private static string letterTemplate = "";
        private List<string> requiredHeadersList = new List<string> { "name", "position", "specialization", "team", "bot name" };

        public static List<Tuple<string, Tuple<List<string>, Dictionary<string, int>>>> ValidationProcess(Logger globalLogger, string originalFileName, string pathToRetrievedFile)
        {
            ValidationWorkflow validator = new ValidationWorkflow();
            return validator.Validation(globalLogger, originalFileName, pathToRetrievedFile);
        }

        List<Tuple<string, Tuple<List<string>, Dictionary<string, int>>>> Validation(Logger globalLogger, string originalFileName, string pathToRetrievedFile)
        {
            globalLogger.Log(0, "Validate input file: " + originalFileName, false);
            string wbName = "";
            string wsName = "";
            try
            {
                Excel.CloseAllExcelInstances();
                Excel.KillExcelProcess();
                Excel.CreateInstance();
                wbName = Excel.OpenWorkbook(pathToRetrievedFile);
            }
            catch
            {
                err = "Error while opening a file. " + "Input file is invalid:" + ListOfReasons + ".";
                letterTemplate = "1";
                globalLogger.Log(1, err, false);
                Excel.CloseAllExcelInstances();
                Excel.KillExcelProcess();
                return null;
            }
            try
            {
                wsName = Excel.GetSheetName(wbName, requiredHeadersList);
            }
            catch (Exception ex)
            {
                err = ex.ToString();
                letterTemplate = "2";
                globalLogger.Log(1, err, false);
                Excel.CloseAllExcelInstances();
                Excel.KillExcelProcess();
                return null;
            }
            if (wsName == "")
            {
                err = "Input file is invalid:" + ListOfReasons + ". Terminating the process.";
                letterTemplate = "2";
                globalLogger.Log(1, err, false);
                Excel.CloseAllExcelInstances();
                Excel.KillExcelProcess();
                return null;
            }
            globalLogger.Log(0, "Successfully found required headers.", false);

            globalLogger.Log(0, "Extracting data from inpute file has been started.", false);
            var data = Excel.GetValuesUnderHeaders(wbName, wsName, requiredHeadersList);
            // Console.WriteLine(data.Count);
            globalLogger.Log(0, "Data has been extracted successfully.", false);

            if (data.Count == 0)
            {
                err = "Input file does not contain data under the required headers; the Bot has stopped processing. Input file: " + originalFileName;
                letterTemplate = "1";
                globalLogger.Log(1, err, false);
                Excel.CloseAllExcelInstances();
                Excel.KillExcelProcess();
                return null;
            }
            Excel.CloseWorkbook(wbName, false);
            Excel.CloseAllExcelInstances();
            globalLogger.Log(0, "Input file was validated successfully. Data has been recieved.", false);

            return data;
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
