using System;
using System.IO;

namespace WordPowerToolsAddin
{
    public static class Logger
    {
        // הנתיב שבו יישמר הלוג: %AppData%\WordPowerToolsAddin\AddinLog.txt
        private static readonly string LogFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WordPowerToolsAddin", "AddinLog.txt");

        public static void LogInfo(string message)
        {
            WriteToFile($"[INFO] {message}");
        }

        public static void LogError(Exception ex, string context = "")
        {
            string contextMessage = string.IsNullOrEmpty(context) ? "" : $" Context: {context} |";
            WriteToFile($"[ERROR]{contextMessage} {ex.Message}{Environment.NewLine}Stack Trace: {ex.StackTrace}");
        }

        private static void WriteToFile(string logMessage)
        {
            try
            {
                string directory = Path.GetDirectoryName(LogFilePath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                
                string formattedMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {logMessage}{Environment.NewLine}";
                File.AppendAllText(LogFilePath, formattedMessage);
            }
            catch
            {
                // השתקה מכוונת: כישלון בכתיבת לוג לעולם לא אמור להקריס את יישום המקור
            }
        }
    }
}