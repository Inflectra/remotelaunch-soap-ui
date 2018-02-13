using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Xml;

namespace Inflectra.RemoteLaunch.Engines.soapUI
{
    /// <summary>
    /// Provides a wrapper class around the SOAP-UI testrunner command-line interface
    /// </summary>
    public class TestRunner
    {
        private string TEST_RUNNER = "testrunner.bat";
        private string LOAD_TEST_RUNNER = "loadtestrunner.bat";

        /// <summary>
        /// The constructor
        /// </summary>
        /// <param name="workingDirectory">The SOAP-UI bin directory (where the testrunner.bat file lives)</param>
        public TestRunner(string workingDirectory)
        {
            this.WorkingDirectory = workingDirectory;
        }

        #region Properties

        /// <summary>
        /// Handle to the application log
        /// </summary>
        public EventLog ApplicationLog
        {
            get;
            set;
        }

        /// <summary>
        /// The test project path
        /// </summary>
        public string ProjectPath
        {
            get;
            set;
        }

        /// <summary>
        /// The test case name
        /// </summary>
        public string TestCase
        {
            get;
            set;
        }

        /// <summary>
        /// The test suite name
        /// </summary>
        public string TestSuite
        {
            get;
            set;
        }

        /// <summary>
        /// Any other command-line switches
        /// </summary>
        public string OtherCommandLineSwitches
        {
            get;
            set;
        }

        /// <summary>
        /// Does this instance of SOAP-UI support the raw 'Data Export' report format
        /// </summary>
        public bool SupportsDataExport
        {
            get;
            set;
        }

        /// <summary>
        /// Do we want to enable trace logging
        /// </summary>
        public bool TraceLogging
        {
            get;
            set;
        }

        /// <summary>
        /// Is this instance of SOAP-UI running a load test
        /// </summary>
        public bool IsLoadTest
        {
            get;
            set;
        }

        /// <summary>
        /// The SOAP-UI working directory
        /// </summary>
        public string WorkingDirectory
        {
            get;
            set;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Executes the current test suite or test case and returns the results
        /// </summary>
        /// <returns>The SOAP-UI results</returns>
        /// <param name="parameters">Any parameters (optional)</param>
        public TestRunnerOutput Execute(Dictionary<string, string> parameters = null)
        {
            if (TraceLogging && ApplicationLog != null)
            {
                ApplicationLog.WriteEntry("SoapUI.TestRunner.Execute: Entering", EventLogEntryType.Information);
            }

            //For SOAP-UI Pro we can use the data export command-line:
            //C:\Program Files\SmartBear\SoapUI-Pro-5.1.1\bin>testrunner.bat -FXML -R"Data Export" -f"C:\Temp\SOAP-UI" -a -s"Requirements Testing" -c"Get Requirements" "C:\Users\Administrator\Documents\SpiraTest-3-0-Web-Service-soapui-project.xml"

            //For SOAP-UI Free Version, we need to use the summary report that's output to the console
            //C:\Program Files\SmartBear\SoapUI-Pro-5.1.1\bin>testrunner.bat -r -a -s"Requirements Testing" -c"Get Requirements" "C:\Users\Administrator\Documents\SpiraTest-3-0-Web-Service-soapui-project.xml"

            //First we need to make sure we have both a test case and test suite
            if (String.IsNullOrEmpty(TestSuite))
            {
                throw new InvalidOperationException("You need to provide a test suite name");
            }
            if (String.IsNullOrEmpty(TestCase))
            {
                throw new InvalidOperationException("You need to provide a test case name");
            }

            //Construct the command line arguments and working folder
            string commandArgs = "";

            //We store the results in a temp output folder
            //If it already exists, delete first to make sure clean
            string outputFolder = Path.Combine(System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Inflectra\\RemoteLaunch_SoapUiRunner");
            if (Directory.Exists(outputFolder))
            {
                Directory.Delete(outputFolder, true);
            }
            Directory.CreateDirectory(outputFolder);

            //Add the standard arguments
            //-a includes all test information, not just errors
            commandArgs += "-a";

            //Add the test suite name
            commandArgs += " -s\"" + TestSuite.Replace("\"", "\\\"") + "\"";

            //Add the test case name
            commandArgs += " -c\"" + TestCase.Replace("\"", "\\\"") + "\"";

            //Add the report format arguments
            if (SupportsDataExport)
            {
                //Save the detailed XML report to a flat file and display the summary report to the console
                commandArgs += " -r -FXML -R\"Data Export\" -f\"" + outputFolder + "\"";
            }
            else
            {
                //Just display the summary report to the console
                commandArgs += " -r";
            }

            //Next we need to add any parameter values
            if (parameters != null && parameters.Count > 0)
            {
                foreach (KeyValuePair<string, string> parameter in parameters)
                {
                    //We only support "Project Properties" currently
                    //Need to remove any spaces or equals signs from the parameter name & value
                    //Also need to quote any quotes
                    string name = parameter.Key.Replace(" ", "").Replace("=", "");
                    string value = parameter.Value.Replace(" ", "").Replace("=", "").Replace("\"", "\"\"");
                    commandArgs += " -P" + name + "=" + value;
                }
            }

            //Any other command-line switches
            if (!String.IsNullOrWhiteSpace(OtherCommandLineSwitches))
            {
                commandArgs += " " + OtherCommandLineSwitches + " ";
            }

            //Next we need to add the path to the test project
            commandArgs += " \"" + ProjectPath + "\"";

            //Finally add on the output redirection
            string consoleOutputFile = Path.Combine(outputFolder, "console.log");
            commandArgs += " > \"" + consoleOutputFile + "\"";

            string runnerBatchFile = TEST_RUNNER;
            if (IsLoadTest)
            {
                runnerBatchFile = LOAD_TEST_RUNNER;
            }

            string filename = Path.Combine(WorkingDirectory, runnerBatchFile);

            //Log the full filename
            if (TraceLogging && ApplicationLog != null)
            {
                ApplicationLog.WriteEntry("SoapUI TestRunner Command Line: " + filename, EventLogEntryType.Information);
                ApplicationLog.WriteEntry("SoapUI TestRunner Command Args: " + commandArgs, EventLogEntryType.Information);
            }

            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = filename;
                startInfo.Arguments = commandArgs;
                startInfo.WorkingDirectory = WorkingDirectory;
                startInfo.UseShellExecute = true;
                startInfo.ErrorDialog = false;
                startInfo.RedirectStandardOutput = false;

                //Now launch the runner, capturing any console output using a pipe
                Process process = new Process();
                process.StartInfo = startInfo;
                process.Start();
                process.WaitForExit();
                process.Close();

                //Read back the console output
                string consoleOutput = "";
                if (File.Exists(consoleOutputFile))
                {
                    consoleOutput = File.ReadAllText(consoleOutputFile);
                }
                else
                {
                    if (ApplicationLog != null)
                    {
                        ApplicationLog.WriteEntry("Unable to find console output log file at: " + consoleOutputFile, EventLogEntryType.Error);
                    }
                }

                //If we have support for the XML data export report, need to open it up as an XML Document
                TestRunnerOutput output;
                if (SupportsDataExport)
                {
                    string exportFile;
                    if (IsLoadTest)
                    {
                        //Test-Suite/Test-Case/LoadTestLog.xml
                        exportFile = Path.Combine(outputFolder, TestSuite.Replace(" ", "-"), TestCase.Replace(" ", "-"), "LoadTestLog.xml");
                    }
                    else
                    {
                        //Test-Suite/Test-Case/TestCaseTestStepResults.xml
                        exportFile = Path.Combine(outputFolder, TestSuite.Replace(" ", "-"), TestCase.Replace(" ", "-"), "TestCaseTestStepResults.xml");
                    }
                    if (File.Exists(exportFile))
                    {
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.Load(exportFile);
                        output = new TestRunnerOutput(consoleOutput, xmlDoc);
                    }
                    else
                    {
                        if (ApplicationLog != null)
                        {
                            ApplicationLog.WriteEntry("Unable to find detailed XML output log file at: " + exportFile, EventLogEntryType.Error);
                        }
                        output = new TestRunnerOutput(consoleOutput);
                    }
                }
                else
                {
                    output = new TestRunnerOutput(consoleOutput);
                }

                if (TraceLogging && ApplicationLog != null)
                {
                    ApplicationLog.WriteEntry("SoapUI.TestRunner.Execute: Exiting", EventLogEntryType.Information);
                }
                return output;
            }
            catch (Exception exception)
            {
                throw new ApplicationException("Unable to launch SOAP-UI TestRunner with arguments " + commandArgs + " in directory " + WorkingDirectory + " (" + exception.Message + ")", exception);
            }
        }

        #endregion
    }
}
