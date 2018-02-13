using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Inflectra.RemoteLaunch.Engines.soapUI
{
    /// <summary>
    /// The output from the test runner
    /// </summary>
    public class TestRunnerOutput
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="consoleOutput"></param>
        public TestRunnerOutput(string consoleOutput)
        {
            this.ConsoleOutput = consoleOutput;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="consoleOutput"></param>
        public TestRunnerOutput(string consoleOutput, XmlDocument xmlOutput)
        {
            this.ConsoleOutput = consoleOutput;
            this.XmlOutput = xmlOutput;
        }

        /// <summary>
        /// The command-line console output
        /// </summary>
        public string ConsoleOutput
        {
            get;
            set;
        }

        /// <summary>
        /// The XML data export report (SOAP-UI pro only)
        /// </summary>
        public XmlDocument XmlOutput
        {
            get;
            set;
        }
    }
}
