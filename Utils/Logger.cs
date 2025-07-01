namespace DocMailer.Utils
{
    /// <summary>
    /// Utilities for logging
    /// </summary>
    public static class Logger
    {
        private static readonly string LogFilePath = Path.Combine("Output", "docmailer.log");

        static Logger()
        {
            // Ensure log directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath) ?? "Output");
        }

        public static void LogInfo(string message)
        {
            Log("INFO", message);
        }

        public static void LogWarning(string message)
        {
            Log("WARNING", message);
        }

        public static void LogError(string message, Exception? exception = null)
        {
            var fullMessage = exception != null ? $"{message} - {exception}" : message;
            Log("ERROR", fullMessage);
        }

        private static void Log(string level, string message)
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            var logEntry = $"[{timestamp}] [{level}] {message}";
            
            // Log to console
            Console.WriteLine(logEntry);
            
            // Log to file
            try
            {
                File.AppendAllText(LogFilePath, logEntry + Environment.NewLine);
            }
            catch
            {
                // Silently ignore logging errors
            }
        }
    }
}
