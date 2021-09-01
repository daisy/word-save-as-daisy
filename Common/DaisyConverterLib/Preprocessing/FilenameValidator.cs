using System.Text.RegularExpressions;

namespace Daisy.SaveAsDAISY.Conversion {
    public class FilenameValidator {
        public Regex AuthorisationPattern { get; set; }
        public string UnauthorizedNameMessage { get; set; }
    }
}
