using System.Text.RegularExpressions;

namespace Daisy.SaveAsDAISY.Conversion
{
    public class StringValidator
    {
        public Regex AuthorisationPattern { get; set; }
        public string UnauthorizedValueMessage { get; set; }


        private static StringValidator _dTBookXMLFileNameFormat = new StringValidator
        {
            AuthorisationPattern = new Regex(@"^[^,]+$"),
            UnauthorizedValueMessage = "Your document file name contains unauthorized characters, that may be automatically replaced by underscores.\r\n" +
                        "Any commas (,) present in the file name should be removed, or they will be replaced by underscores automatically."
        };
        public static StringValidator DTBookXMLFileNameFormat { get => _dTBookXMLFileNameFormat; }


        private static StringValidator _DAISYFileNameFormat = new StringValidator
        {
            AuthorisationPattern = new Regex(@"^[a-zA-Z0-9_\-\.]+$"),
            UnauthorizedValueMessage = "Your document file name contains unauthorized characters, that may be automatically replaced by underscores.\r\n" +
                        "Only Alphanumerical letters (a-z, A-Z, 0-9), hyphens (-), dots (.) " +
                            "and underscores (_) are allowed in DAISY file names." +
                        "\r\nAny other characters (including spaces) will be replaced automaticaly by underscores."
        };
        public static StringValidator DAISYFileNameFormat { get => _DAISYFileNameFormat; }
    }
}
