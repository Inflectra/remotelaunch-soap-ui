using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

using Inflectra.RemoteLaunch.Interfaces;
using Inflectra.RemoteLaunch.Interfaces.DataObjects;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;

namespace Inflectra.RemoteLaunch.Engines.soapUI
{
    /// <summary>
    /// Implements the IAutomationEngine class for integration with SmartBear SOAP-UI
    /// This class is instantiated by the RemoteLaunch application
    /// </summary>
    /// <remarks>
    /// The AutomationEngine class provides some of the generic functionality
    /// </remarks>
    public class SoapUIEngine : AutomationEngine, IAutomationEngine4
    {
        private const string CLASS_NAME = "SoapUIEngine";

        private const string AUTOMATION_ENGINE_TOKEN = "SoapUI";
        private const string AUTOMATION_ENGINE_VERSION = "4.0.1";

        /// <summary>
        /// Constructor
        /// </summary>
        public SoapUIEngine()
        {
            //Set status to OK
            base.status = EngineStatus.OK;
        }

        /// <summary>
        /// Returns the author of the test automation engine
        /// </summary>
        public override string ExtensionAuthor
        {
            get
            {
                return "Inflectra Corporation";
            }
        }

        /// <summary>
        /// The unique GUID that defines this automation engine
        /// </summary>
        public override Guid ExtensionID
        {
            get
            {
                return new Guid("{6DCE96C2-33E8-42D9-ABD9-93BF1A0896E2}");
            }
        }

        /// <summary>
        /// Returns the display name of the automation engine
        /// </summary>
        public override string ExtensionName
        {
            get
            {
                return "SOAP-UI Automation Engine";
            }
        }

        /// <summary>
        /// Returns the unique token that identifies this automation engine to SpiraTest
        /// </summary>
        public override string ExtensionToken
        {
            get
            {
                return AUTOMATION_ENGINE_TOKEN;
            }
        }

        /// <summary>
        /// Returns the version number of this extension
        /// </summary>
        public override string ExtensionVersion
        {
            get
            {
                return AUTOMATION_ENGINE_VERSION;
            }
        }

        /// <summary>
        /// Adds a custom settings panel for configuring the SOAP-UI environment
        /// </summary>
        public override System.Windows.UIElement SettingsPanel
        {
            get
            {
                return new SoapUISettings();
            }
            set
            {
                SoapUISettings soapUISettings = (SoapUISettings)value;
                soapUISettings.SaveSettings();
            }
        }

        public override AutomatedTestRun StartExecution(AutomatedTestRun automatedTestRun)
        {
            //Not used since we implement the V4 API instead
            throw new NotImplementedException();
        }

        /// <summary>
        /// This is the main method that is used to start automated test execution
        /// </summary>
        /// <param name="automatedTestRun">The automated test run object</param>
        /// <returns>Either the populated test run or an exception</returns>
        /// <param name="projectId">The id of the project</param>
        public AutomatedTestRun4 StartExecution(AutomatedTestRun4 automatedTestRun, int projectId)
        {
            //Set status to OK
            base.status = EngineStatus.OK;

            try
            {
                if (Properties.Settings.Default.TraceLogging && applicationLog != null)
                {
                    applicationLog.WriteEntry("SoapUIEngine.StartExecution: Entering", EventLogEntryType.Information);
                }

                if (automatedTestRun == null)
                {
                    throw new InvalidOperationException("The automatedTestRun object provided was null");
                }

                //Instantiate the SOAP-UI runner wrapper class
                TestRunner soapUiRunner = new TestRunner(Properties.Settings.Default.Location);

                //Set the license type (pro vs. free)
                soapUiRunner.SupportsDataExport = Properties.Settings.Default.ProLicense;

                //Specify if this is a load test or not
                soapUiRunner.IsLoadTest = Properties.Settings.Default.LoadTest;

                if (Properties.Settings.Default.TraceLogging && applicationLog != null)
                {
                    LogEvent("Starting test execution", EventLogEntryType.Information);                    
                }
                soapUiRunner.TraceLogging = Properties.Settings.Default.TraceLogging;

                //Pass the application log handle
                soapUiRunner.ApplicationLog = this.applicationLog;

                //See if we have an attached or linked test script
                //For squish we only support linked test cases
                if (automatedTestRun.Type == AutomatedTestRun4.AttachmentType.URL)
                {
                    //The "URL" of the test is a combination of project filename, project suite name and test case name:
                    //Project File Name|Test Suite Name|Test Case Name
                    //e.g. [MyDocuments]\SpiraTest-3-0-Web-Service-soapui-project.xml|Requirements Testing|Get Requirements

                    //See if we have any pipes in the 'filename' that include additional options
                    string[] filenameElements = automatedTestRun.FilenameOrUrl.Split('|');

                    //Make sure we have all three elements (the fourth is optional)
                    if (filenameElements.Length < 3)
                    {
                        throw new ArgumentException(String.Format("You need to provide a project file, test suite and test case name separated by pipe (|) characters. Only {0} elements were provided.", filenameElements.Length));
                    }

                    //To make it easier, we have certain shortcuts that can be used in the path
                    string path = filenameElements[0];
                    path = path.Replace("[MyDocuments]", Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments));
                    path = path.Replace("[CommonDocuments]", Environment.GetFolderPath(System.Environment.SpecialFolder.CommonDocuments));
                    path = path.Replace("[DesktopDirectory]", Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory));
                    path = path.Replace("[ProgramFiles]", Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFiles));
                    path = path.Replace("[ProgramFilesX86]", Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFilesX86));

                    //First make sure that the File exists, or folder if it's a composite project
                    if (File.Exists(path) || Directory.Exists(path))
                    {
                        //Set the path, test suite and test case on the runner
                        soapUiRunner.ProjectPath = path;
                        soapUiRunner.TestSuite = filenameElements[1];
                        soapUiRunner.TestCase = filenameElements[2];
                        if (filenameElements.Length > 3)
                        {
                            soapUiRunner.OtherCommandLineSwitches = filenameElements[3];
                        }
                    }
                    else
                    {
                        throw new ArgumentException(String.Format("The provided project filepath '{0}' does not exist on the host!", path ));
                    }
                }
                else
                {
                    //We have an embedded script which we need to execute directly
                    //This is not currently supported since SOAP-UI uses XML files which cannot be easily edited in SpiraTest
                    throw new InvalidOperationException("The SOAP-UI automation engine only supports linked test scripts");
                }

                //Actually run the test
                DateTime startDate = DateTime.Now;
                TestRunnerOutput output;
                //See if we have any parameters we need to pass
                if (automatedTestRun.Parameters == null)
                {
                    if (Properties.Settings.Default.TraceLogging && applicationLog != null)
                    {
                        LogEvent("Test Run has no parameters", EventLogEntryType.Information);
                    }

                    //Run the test
                    output = soapUiRunner.Execute();
                }
                else
                {
                    if (Properties.Settings.Default.TraceLogging && applicationLog != null)
                    {
                        LogEvent("Test Run has parameters", EventLogEntryType.Information);
                    }

                    Dictionary<string, string> parameters = new Dictionary<string, string>();
                    foreach (TestRunParameter testRunParameter in automatedTestRun.Parameters)
                    {
                        string parameterName = testRunParameter.Name.ToLowerInvariant();
                        if (!parameters.ContainsKey(parameterName))
                        {
                            //Make sure the parameters are lower case for comparing later
                            if (Properties.Settings.Default.TraceLogging && applicationLog != null)
                            {
                                LogEvent("Adding test run parameter " + parameterName + " = " + testRunParameter.Value, EventLogEntryType.Information);
                            }
                            parameters.Add(parameterName, testRunParameter.Value);
                        }
                    }

                    //Run the test with parameters
                    output = soapUiRunner.Execute(parameters);
                }
                DateTime endDate = DateTime.Now;

                //Specify the start/end dates
                automatedTestRun.StartDate = startDate;
                automatedTestRun.EndDate = endDate;
                automatedTestRun.RunnerTestName = soapUiRunner.TestSuite + " / " + soapUiRunner.TestCase;

                //Need to parse the results

                //SoapUI 3.6.1 TestCaseRunner Summary
                //-----------------------------
                //Time Taken: 837ms
                //Total TestSuites: 0
                //Total TestCases: 1 (0 failed)
                //Total TestSteps: 3
                //Total Request Assertions: 6
                //Total Failed Assertions: 0
                //Total Exported Results: 3

                //Parse the summary result log
                string[] lines = output.ConsoleOutput.Split('\n');
                int stepsCount = 0;
                int assertions = 0;
                int failedAssertions = 0;
                int failedTestCases = 0;
                foreach (string line in lines)
                {
                    //Extract the various counts
                    if (line.StartsWith("Total TestSteps:"))
                    {
                        string value = line.Substring("Total TestSteps:".Length).Trim();
                        int intValue;
                        if (Int32.TryParse(value, out intValue))
                        {
                            stepsCount = intValue;
                        }
                    }
                    if (line.StartsWith("Total Request Assertions:"))
                    {
                        string value = line.Substring("Total Request Assertions:".Length).Trim();
                        int intValue;
                        if (Int32.TryParse(value, out intValue))
                        {
                            assertions = intValue;
                        }
                    }
                    if (line.StartsWith("Total Failed Assertions:"))
                    {
                        string value = line.Substring("Total Failed Assertions:".Length).Trim();
                        int intValue;
                        if (Int32.TryParse(value, out intValue))
                        {
                            failedAssertions = intValue;
                        }
                    }

                    //Use Regex to parse the number of failed test cases
                    Regex regex = new Regex(@"^Total TestCases: (\d+) \((\d) failed\)");
                    if (regex.IsMatch(line))
                    {
                        Match match = regex.Match(line);
                        if (match != null && match.Groups.Count >= 3)
                        {
                            int intValue;
                            if (Int32.TryParse(match.Groups[2].Value, out intValue))
                            {
                                failedTestCases = intValue;
                            }
                        }
                    }
                }

                automatedTestRun.ExecutionStatus = (failedAssertions > 0 || failedTestCases > 0) ? AutomatedTestRun4.TestStatusEnum.Failed : AutomatedTestRun4.TestStatusEnum.Passed;
                automatedTestRun.RunnerMessage = String.Format("{0} test steps completed with {1} request assertions, {2} failed assertions and {3} failed test cases.", stepsCount, assertions, failedAssertions, failedTestCases);
                automatedTestRun.RunnerStackTrace = output.ConsoleOutput;
                automatedTestRun.RunnerAssertCount = failedAssertions;

                //If we have an instance of SOAP-UI Pro, there are more detailed XML reports available
                //
                //SOAP-UI TEST
                //<TestCaseTestStepResults>
                //  <result>
                //    <message>Step 0 [Authenticate] OK: took 617 ms</message> 
                //    <name>Authenticate</name> 
                //    <order>1</order> 
                //    <started>23:03:30.871</started> 
                //    <status>OK</status> 
                //    <timeTaken>617</timeTaken> 
                //  </result>
                //</TestCaseTestStepResults>
                //
                //LOAD-UI TEST
                //<LoadTestLog>
                //<entry>
                //  <discarded>false</discarded>
                //  <error>false</error>
                //  <message>LoadTest started at Tue Jun 24 11:31:35 EST 2014</message>
                //  <timeStamp>1403573495035</timeStamp>
                //  <type>Message</type>
                //</entry>
                //<entry>
                //  <discarded>false</discarded>
                //  <error>true</error>
                //  <message><![CDATA[TestStep [tc13_listAccountBalanceHistoryV1] result status is FAILED; [[RegExp assertion (from spreadsheet)] assert strResponse.contains("<ClientID>")
                //     |           |
                //     |           false
                //     <?xml version="1.0" encoding="UTF-8"?><SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/"><SOAP-ENV:Body><ns0:ListAccountBalanceHistoryResponseV1 xmlns:ns0="http://www.qsuper.com.au/services/business/AccountEnquiryV1"><ns0:Status>Success</ns0:Status><ns0:ResponsePayload><ns0:Client><ClientID xmlns="">194919272</ClientID></ns0:Client></ns0:ResponsePayload></ns0:ListAccountBalanceHistoryResponseV1></SOAP-ENV:Body></SOAP-ENV:Envelope>] [threadIndex=0]]]></message>
                //  <targetStepName>tc13_listAccountBalanceHistoryV1</targetStepName>
                //  <timeStamp>1403573495573</timeStamp>
                //  <type>Step Status</type>
                //</entry>
                //</LoadTestLog>
                if (Properties.Settings.Default.ProLicense)
                {
                    if (output.XmlOutput != null)
                    {
                        bool errorFound = false;    //Sometimes failure will be in the steps
                        automatedTestRun.TestRunSteps = new List<TestRunStep4>();
                        int position = 1;

                        //See if we have a load test or not
                        if (Properties.Settings.Default.LoadTest)
                        {
                            XmlNodeList xmlNodes = output.XmlOutput.SelectNodes("LoadTestLog/entry");
                            foreach (XmlNode xmlNode in xmlNodes)
                            {
                                //Add the message to the stack-trace
                                string message = xmlNode.SelectSingleNode("message").InnerText;
                                string type = xmlNode.SelectSingleNode("type").InnerText;
                                string targetStepName = "";
                                if (xmlNode.SelectSingleNode("targetStepName") != null)
                                {
                                    targetStepName = xmlNode.SelectSingleNode("targetStepName").InnerText;
                                }
                                string error = xmlNode.SelectSingleNode("error").InnerText;

                                //Add the 'test step'
                                TestRunStep4 testRunStep = new TestRunStep4();
                                testRunStep.ExecutionStatusId = (error == "true") ? (int)AutomatedTestRun4.TestStatusEnum.Failed : (int)AutomatedTestRun4.TestStatusEnum.Passed;
                                testRunStep.Description = type + ": " + targetStepName;
                                testRunStep.ActualResult = message;
                                testRunStep.Position = position++;
                                automatedTestRun.TestRunSteps.Add(testRunStep);
                                if (error == "true")
                                {
                                    errorFound = true;
                                }
                            }
                        }
                        else
                        {
                            XmlNodeList xmlNodes = output.XmlOutput.SelectNodes("TestCaseTestStepResults/result");
                            foreach (XmlNode xmlNode in xmlNodes)
                            {
                                //Get the message
                                string message = xmlNode.SelectSingleNode("message").InnerText;

                                //Add the 'test step'
                                string status = xmlNode.SelectSingleNode("status").InnerText;
                                TestRunStep4 testRunStep = new TestRunStep4();
                                switch (status)
                                {
                                    case "OK":
                                        testRunStep.ExecutionStatusId = (int)AutomatedTestRun4.TestStatusEnum.Passed;
                                        break;

                                    default:
                                        testRunStep.ExecutionStatusId = (int)AutomatedTestRun4.TestStatusEnum.Failed;
                                        errorFound = true;
                                        break;
                                }
                                testRunStep.Description = message;
                                testRunStep.ActualResult = status;
                                testRunStep.Position = position++;
                                automatedTestRun.TestRunSteps.Add(testRunStep);
                            }
                        }

                        if (errorFound && (automatedTestRun.ExecutionStatus == AutomatedTestRun4.TestStatusEnum.Passed || automatedTestRun.ExecutionStatus == AutomatedTestRun4.TestStatusEnum.NotRun || automatedTestRun.ExecutionStatus == AutomatedTestRun4.TestStatusEnum.Caution))
                        {
                            automatedTestRun.ExecutionStatus = AutomatedTestRun4.TestStatusEnum.Failed;
                        }
                    }
                    else
                    {
                        if (applicationLog != null)
                        {
                            LogEvent("Unable to access the SOAP-UI Pro Detailed XML Log File", EventLogEntryType.Error);
                        }
                    }
                }

                if (Properties.Settings.Default.TraceLogging && applicationLog != null)
                {
                    applicationLog.WriteEntry("SoapUIEngine.StartExecution: Entering", EventLogEntryType.Information);
                }

                //Report as complete               
                base.status = EngineStatus.OK;
                return automatedTestRun;
            }
            catch (Exception exception)
            {
                //Log the error and denote failure
                LogEvent(exception.Message + " (" + exception.StackTrace + ")", EventLogEntryType.Error);

                //Report as completed with error
                base.status = EngineStatus.Error;
                throw exception;
            }
        }

        /// <summary>
        /// Returns the full token of a test caseparameter from its name
        /// </summary>
        /// <param name="parameterName">The name of the parameter</param>
        /// <returns>The tokenized representation of the parameter used for search/replace</returns>
        /// <remarks>We use the same parameter format as Ant/NAnt</remarks>
        public static string CreateParameterToken(string parameterName)
        {
            return "${" + parameterName + "}";
        }
    }
}
