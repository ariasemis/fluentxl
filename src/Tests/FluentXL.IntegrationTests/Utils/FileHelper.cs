using System;
using System.IO;

namespace FluentXL.IntegrationTests.Utils
{
    internal static class FileHelper
    {
        public static string GetOutputDirectory()
            => Path.Combine(Environment.CurrentDirectory, @"Output\");

        public static bool IsFileLocked(string filePath)
            => IsFileLocked(new FileInfo(filePath));

        public static bool IsFileLocked(FileInfo file)
        {
            if (!file.Exists)
                return false;

            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            return false;
        }
    }
}
