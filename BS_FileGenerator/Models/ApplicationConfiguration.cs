namespace BS_FileGenerator.Models
{
    public class ApplicationConfiguration
    {
        public bool IsLoggingEnabled { set; get; } = false;
        public string LoggingFilePath { set; get; }
    }
}
