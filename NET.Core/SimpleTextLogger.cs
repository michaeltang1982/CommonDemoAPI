using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sierra.NET.Core;

namespace Sierra.NET.Core
{
    /// <summary>
    /// simple implementation of ILogger which logs messages to an internal list of strings
    /// </summary>
    public class SimpleTextLogger: ILogger
    {
        public List<string> Messages { get; private set; }

        public SimpleTextLogger() { Initialise(); }

        public void LogError(Exception ex) { this.Messages.Add(ex.Message); }
        public void LogError(Exception ex, string message) { this.Messages.Add(message); }
        public void LogError(string message, string stackTrace) { this.Messages.Add(message); }
        public void LogWarning(string message) { this.Messages.Add("WARNING:" + message); }
        public void LogInformation(string message) { this.Messages.Add("INFO:" + message); }
        public void LogVerbose(string message) { this.Messages.Add(message); }
        public void LogVerbose(string message, string methodName) { this.Messages.Add(message); }


        private void Initialise() { this.Messages = new List<string>(); }
    }
}
