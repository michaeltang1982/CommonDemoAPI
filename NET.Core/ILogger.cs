using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sierra.NET.Core
{
    public interface ILogger
    {
        void LogError(Exception ex);
        void LogError(Exception ex, string message);
        void LogError(string message, string stackTrace);
        void LogWarning(string message);
        void LogInformation(string message);
        void LogVerbose(string message);
        void LogVerbose(string message, string methodName);

    }
    
}
