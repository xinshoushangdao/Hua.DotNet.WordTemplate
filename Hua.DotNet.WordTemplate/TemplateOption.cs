using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hua.DotNet.WordTemplate
{

    /// <summary>
    /// 张三 Demo
    /// </summary>

    public class TemplateOption
    {
        public string Pattern => @"\$\{[^\}]*\}";

        public string StartPattern { get; set; } = "${";

        public string EndPattern { get; set; } = "}";

        public string Separator = ":";

        public string StartTag = "Start";

        public string EndTag = "End";

    }
}
