using System;

namespace Converter.Parsing
{
    /// <summary>
    /// Простой логгер для отладочных сообщений
    /// </summary>
    public static class Logger
    {
        /// <summary>
        /// Логирует информационное сообщение
        /// </summary>
        public static void Info(string message)
        {
            System.Diagnostics.Debug.WriteLine($"[INFO] {DateTime.Now:HH:mm:ss} {message}");
            Console.WriteLine($"[INFO] {DateTime.Now:HH:mm:ss} {message}");
        }

        /// <summary>
        /// Логирует сообщение об ошибке
        /// </summary>
        public static void Error(Exception ex, string context = "")
        {
            var message = string.IsNullOrEmpty(context) 
                ? $"Error: {ex.Message}" 
                : $"Error in {context}: {ex.Message}";
            
            System.Diagnostics.Debug.WriteLine($"[ERROR] {DateTime.Now:HH:mm:ss} {message}");
            Console.WriteLine($"[ERROR] {DateTime.Now:HH:mm:ss} {message}");
        }
    }
}
