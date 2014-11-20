using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WordToExcel
{
    internal class Logger
    {
        private static NLog.Logger fLog = null;

        private static NLog.Logger Log
        {
            get
            {
                if (fLog == null) CreateLogger();
                return fLog;
            }
        }

        private static string getThreadName()
        {
            try
            {
                string thdName = Thread.CurrentThread.Name;
                if (string.IsNullOrEmpty(thdName))
                {
                    thdName = "Thread_" + Thread.CurrentThread.ManagedThreadId;
                }
                return thdName;
            }
            catch
            {
                return "Thread_unknown";
            }
        }

        public static void SetCurrentThreadName(string prefix)
        {
            string tmp = prefix;
            if (!tmp.EndsWith("_")) tmp += "_";
            string threadName = tmp + Thread.CurrentThread.ManagedThreadId;
            try
            {
                Thread.CurrentThread.Name = threadName;
            }
            catch (Exception ex)
            {
                if (Thread.CurrentThread.Name != threadName)
                {
                    Log.Error(getThreadName() + ": " + "Thread cannot be renamed! Current name - " + Thread.CurrentThread.Name + ", New name - " + threadName);
                    Log.Error(getThreadName() + ": " + "Exception: " + ex.Message + ", StackTrace: " + ex.StackTrace);
                    //ProcessException.DoProcess(new Exception("Thread cannot be renamed! Current name - " + Thread.CurrentThread.Name + ", New name - " + threadName, ex));
                }
            }
        }

        private static void CreateLogger()
        {
            try
            {
                fLog = LogManager.GetCurrentClassLogger();
                string assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                NLog.LogManager.Configuration = new NLog.Config.XmlLoggingConfiguration(assemblyFolder + "\\NLog.config", true);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError("Error creating logger: " + ex.Message + "\nStackTrace: " + ex.StackTrace);
            }
        }

        public static void LogE(string msg, Exception ex)
        {
            msg = string.Format(msg + "\nMessage: {0}\nStackTrace: {1}", ex.Message, ex.StackTrace);
            Log.Error(getThreadName() + ": " + msg);
            System.Diagnostics.Trace.TraceError(getThreadName() + ": " + msg);
        }

        public static void LogI(string msg)
        {
            Log.Info(getThreadName() + ": " + msg);
            System.Diagnostics.Trace.TraceInformation(getThreadName() + ": " + msg);
        }

        public static void LogW(string msg)
        {
            Log.Info(getThreadName() + ": " + msg);
            System.Diagnostics.Trace.TraceWarning(getThreadName() + ": " + msg);
        }
    }
}
