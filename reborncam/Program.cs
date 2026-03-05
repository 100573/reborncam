using System;
using System.IO;
using System.Windows.Forms;

namespace reborncam
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Disable MSMF hardware transforms in OpenCV which can cause unstable behavior with some cameras
            Environment.SetEnvironmentVariable("OPENCV_VIDEOIO_MSMF_ENABLE_HW_TRANSFORMS", "0");

            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();

            // global exception handlers to capture startup/runtime errors
            AppDomain.CurrentDomain.UnhandledException += (s, e) => LogAndShowUnhandled(e.ExceptionObject as Exception, "AppDomain.UnhandledException");
            Application.ThreadException += (s, e) => LogAndShowUnhandled(e.Exception, "Application.ThreadException");

            try
            {
                Application.Run(new Form1());
            }
            catch (Exception ex)
            {
                LogAndShowUnhandled(ex, "Application.Run.Exception");
                throw;
            }
        }

        private static void LogAndShowUnhandled(Exception? ex, string kind)
        {
            try
            {
                var logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log");
                Directory.CreateDirectory(logDir);
                var path = Path.Combine(logDir, "diagnostics.log");
                var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} | {kind} | {ex?.ToString() ?? "(null)"}";
                File.AppendAllLines(path, new[] { line });
            }
            catch { }

            try
            {
                MessageBox.Show($"Unhandled error ({kind}): {ex?.Message}\nSee log\\diagnostics.log for details.", "URANUS - Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch { }
        }
    }
}