namespace BS_FileGenerator.Models
{
    public class PasswordSettings
    {
        public int PasswordMinLength { set; get; }
        public int PasswordMaxLength { set; get; }
        public int PasswordRegexFormula { set; get; }
        public int TemporaryPasswordExpiryTimeInSeconds { set; get; }        
    }
}
