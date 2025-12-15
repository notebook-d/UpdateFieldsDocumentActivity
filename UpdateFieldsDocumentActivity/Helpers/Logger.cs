using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UpdateFieldsDocumentActivity.Helpers
{
    internal static class Logger
    {
        private static readonly object _lock = new object();
        private static string _logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ASCON", "Pilot-ICE Enterprise", "Logs", "Units Logs");
        private static readonly string _logFileNamePrefix = "UpdateFieldsDocumentActivity";
        private static string _currentLogFilePath;

        // Признак ведения лога (по умолчанию включен)
        private static bool _isLoggingEnabled = true;

        public static string LogDirectory
        {
            get => _logDirectory;
            set => _logDirectory = value ?? Path.GetTempPath();
        }

        /// <summary>
        /// Включено ли ведение лога
        /// </summary>
        public static bool IsLoggingEnabled
        {
            get => _isLoggingEnabled;
            set => _isLoggingEnabled = value;
        }

        public static void WriteLog(LogEntry logEntry)
        {
            // Проверяем, включено ли ведение лога
            if (!_isLoggingEnabled)
                return;

            try
            {
                string logFilePath = GetLogFilePath();
                string str = logEntry.ToString();
                lock (_lock)
                {
                    EnsureDirectoryExists(logFilePath);
                    File.AppendAllText(logFilePath, str + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public static async Task WriteLogAsync(LogEntry logEntry)
        {
            // Проверяем, включено ли ведение лога
            if (!_isLoggingEnabled)
                return;

            try
            {
                string logFilePath = GetLogFilePath();
                string logMessage = logEntry.ToString();
                await Task.Run(() =>
                {
                    lock (_lock)
                    {
                        EnsureDirectoryExists(logFilePath);
                        File.AppendAllText(logFilePath, logMessage + Environment.NewLine, Encoding.UTF8);
                    }
                });
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public static string GetLogFilePath()
        {
            if (!string.IsNullOrEmpty(_currentLogFilePath))
                return _currentLogFilePath;
            _currentLogFilePath = Path.Combine(LogDirectory, _logFileNamePrefix + ".log");
            return _currentLogFilePath;
        }

        private static void EnsureDirectoryExists(string filePath)
        {
            string directoryName = Path.GetDirectoryName(filePath);
            if (string.IsNullOrEmpty(directoryName) || Directory.Exists(directoryName))
                return;
            Directory.CreateDirectory(directoryName);
        }

        /// <summary>
        /// Включить ведение лога
        /// </summary>
        public static void EnableLogging()
        {
            _isLoggingEnabled = true;
        }

        /// <summary>
        /// Выключить ведение лога
        /// </summary>
        public static void DisableLogging()
        {
            _isLoggingEnabled = false;
        }
    }
}
