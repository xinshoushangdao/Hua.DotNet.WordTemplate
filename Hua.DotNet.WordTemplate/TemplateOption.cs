using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hua.DotNet.WordTemplate
{
    public class TemplateOption
    {
        public string Pattern
        {
            get
            {
                return @"\$\{[^\}]*\}";
            }
        }

        public string StartPattern { get; set; } = "${";

        public string EndPattern { get; set; } = "}";

        public string StartTag = "<";

        public string EndTag = ">";

    }
}
