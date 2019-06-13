using Serilog;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DQILogger.Core
{
    public static class Logger
    {
        private static readonly ILogger _perfLogger;
        private static readonly ILogger _usageLogger;
        private static readonly ILogger _errorLogger;
        private static readonly ILogger _diagnosticLogger;
        static Logger()
        {
            _perfLogger = new LoggerConfiguration()
                .WriteTo.File(path: "C:\\users\\ctaylor\\Source\\AppLogs\\perf.txt")
                .CreateLogger();

            _usageLogger = new LoggerConfiguration()
                .WriteTo.File(path: "C:\\users\\ctaylor\\Source\\AppLogs\\usage.txt")
                .CreateLogger();

            _errorLogger = new LoggerConfiguration()
                .WriteTo.File(path: "C:\\users\\ctaylor\\Source\\AppLOgs\\error.txt")
                .CreateLogger();

            _diagnosticLogger = new LoggerConfiguration()
                .WriteTo.File(path: "C:\\users\\ctaylor\\Source\\AppLogs\\diagnostic.txt")
                .CreateLogger();
        }
        //Public Methods used by GetProjectData
        public static void WritePerf(LogDetail infoToLog)
        {
            _perfLogger.Write(LogEventLevel.Information, "{@LogDetail}", infoToLog);
        }
        public static void WriteUsage(LogDetail infoToLog)
        {
            _usageLogger.Write(LogEventLevel.Information, "{@LogDetail}", infoToLog);
        }
        public static void WriteError(LogDetail infoToLog)
        {
            /*
            if (infoToLog.Exception != null)
            {
                var procName = FindProcName(infoToLog.Exception);
                infoToLog.Location = string.IsNullOrEmpty(procName) ? infoToLog.Location : procName;
                infoToLog.Message = GetMessageFromException(infoToLog.Exception);
            }
            */
            _errorLogger.Write(LogEventLevel.Information, "{@LogDetail}", infoToLog);
        }
        public static void WriteDiagnostic(LogDetail infoToLog)
        {// Switch in Config file
            var writeDiagnostics = Convert.ToBoolean(ConfigurationManager.AppSettings["EnableDiagnostics"]);
            if (!writeDiagnostics)
                return;

            _diagnosticLogger.Write(LogEventLevel.Information, "{@LogDetail}", infoToLog);
        }

        private static string FindProcName(Exception ex)
        {
            var sqlEx = ex as SqlException;
            if (sqlEx != null)
            {
                var procName = sqlEx.Procedure;
                if (!string.IsNullOrEmpty(procName))
                    return procName;
            }

            if (!string.IsNullOrEmpty((string)ex.Data["Procedure"]))
            {
                return (string)ex.Data["Procedure"];
            }

            if (ex.InnerException != null)
                return FindProcName(ex.InnerException);

            return null;
        }

    }
}
