using System;
using System.Data;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace CSharpAddIn
{
    public class LogHelper
    {
        private LogHelper(){}

        public static readonly Log4Net.ILog logdebug = Log4Net.LogManager.GetLogger("logdebug");
        public static readonly Log4Net.ILog loginfo = Log4Net.LogManager.GetLogger("loginfo");
        public static readonly Log4Net.ILog logwarn = Log4Net.LogManager.GetLogger("logwarn");
        public static readonly Log4Net.ILog logerror = Log4Net.LogManager.GetLogger("logerror");
        public static readonly Log4Net.ILog logfatal = Log4Net.LogManager.GetLogger("logfatal");

        public static void SetConfig()
        {
            Log4Net.Config.XmlConfigurator.Configure();
        }

        public static void SetConfig(FileInfo configFile)
        {
            Log4Net.Config.XmlConfigurator.Configure(configFile);
        }

        public static void LogDebug(string debug)
        {
            if (logdebug.IsDebugEnabled)
                logdebug.Debug(debug);
        }

        public static void LogDebug(string debug, Exception ex)
        {
            if (logdebug.IsDebugEnabled)
                logdebug.Debug(debug, ex);
        }

        public static void LogInfo(string info)
        {
            if (loginfo.IsInfoEnabled)
                loginfo.Info(info);
        }

        public static void LogInfo(string info, Exception ex)
        {
            if (loginfo.IsInfoEnabled)
                loginfo.Error(info, ex);
        }

        public static void LogWarn(string warn)
        {
            if (logwarn.IsWarnEnabled)
                logwarn.Info(warn);
        }

        public static void LogWarn(string warn, Exception ex)
        {
            if (logwarn.IsWarnEnabled)
                logwarn.Warn(warn, ex);
        }

        public static void LogError(string error)
        {
            if (logerror.IsErrorEnabled)
                logerror.Info(error);
        }

        public static void LogError(string error, Exception ex)
        {
            if (logerror.IsErrorEnabled)
                logerror.Error(error, ex);
        }

        public static void LogFatal(string fatal)
        {
            if (logfatal.IsFatalEnabled)
                logfatal.Fatal(fatal);
        }

        public static void LogFatal(string fatal, Exception ex)
        {
            if (logfatal.IsFatalEnabled)
                logfatal.Fatal(fatal, ex);
        }
    }
}
