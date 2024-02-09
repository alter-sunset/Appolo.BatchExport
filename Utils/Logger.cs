using System;
using System.IO;
using System.Security.Policy;
using System.Windows.Forms;

namespace Appolo.Utilities
{
    public class Logger : IDisposable
    {
        public string Path { get; set; }
        private DateTime StartTime { get; set; }
        private string FileName { get; set; }
        private string FilePath { get; set; }
        private const string BaseError = "Произошла срань.";
        public int ErrorCount { get; set; }
        public int SuccessCount { get; set; }

        public Logger(string path)
        {
            Path = path;
            StartTime = DateTime.Now;
            ErrorCount = 0;
            SuccessCount = 0;
            FileName = $"Log_{StartTime:yy-MM-dd_HH-mm-ss}.log";
            FilePath = $@"{Path}\{FileName}";

            string lineToWrite = $"Initial launch at {StartTime}.";
            File.WriteAllText(FilePath, lineToWrite);
        }

        public void Error(string error, Exception ex)
        {
            string lineToWrite = $"Error at {DateTime.Now}. {BaseError} {error} {ex.Message}";

            LineWriter(lineToWrite);

            ErrorCount++;
        }

        public void Error(string error)
        {
            string lineToWrite = $"Error at {DateTime.Now}. {BaseError} {error}";

            try
            {
                LineWriter(lineToWrite);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            ErrorCount++;
        }

        public void Success(string success)
        {
            string lineToWrite = $"Success at {DateTime.Now}. {success}";

            LineWriter(lineToWrite);

            SuccessCount++;
        }

        public void FileOpened()
        {
            const string lineToWrite = "File succesfully opened.";

            LineWriter(lineToWrite);
        }

        public void Start(string file)
        {
            string lineToWrite = $"Started work at {DateTime.Now} on {file}";

            LineWriter(lineToWrite);
        }

        public void LineBreak()
        {
            const string lineToWrite = "--||--";

            LineWriter(lineToWrite);
        }
        public void TimeForFile(DateTime startTime)
        {
            TimeSpan timeSpan = DateTime.Now - startTime;
            string lineToWrite = $"Time spent for file {timeSpan}";

            LineWriter(lineToWrite);
        }

        public void TimeTotal()
        {
            TimeSpan timeSpan = DateTime.Now - StartTime;
            string lineToWrite = $"Total time spent {timeSpan}";

            LineWriter(lineToWrite);
        }

        public void Hash(string hash)
        {
            string lineToWrite = $"MD5 Hash of file is: {hash} at {DateTime.Now}";

            LineWriter(lineToWrite);
        }

        public void ErrorTotal()
        {
            string lineToWrite = $"Done! There were {ErrorCount} errors out of {ErrorCount + SuccessCount} files.";

            LineWriter(lineToWrite);
        }

        private void LineWriter(string lineToWrite)
        {
            try
            {
                if (!File.Exists(FilePath))
                {
                    File.WriteAllText(FilePath, lineToWrite);
                }
                else
                {
                    string toWrite = "\n" + lineToWrite;
                    File.AppendAllText(FilePath, toWrite);
                }
            }
            catch
            {
                MessageBox.Show("Проблемы с файлом логов");
            }
        }

        public void Dispose()
        {

        }
    }
}
