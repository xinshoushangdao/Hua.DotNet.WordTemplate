using NPOI.POIFS.Crypt.Dsig;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Hua.DotNet.WordTemplate
{
    /// <summary>
    /// Word template 
    /// </summary>
    public class WordTemplate<T>
    {
        private readonly string _templatePath;
        private readonly T _data;
        private readonly TemplateOption _option;
        private int _desIndex = 0;
        delegate void Export<T>(T value);

        public WordTemplate(T data,string templatePath, TemplateOption? option=null)
        {
            this._templatePath = templatePath;
            this._data = data;
            if (option==null)
            {
                this._option = new TemplateOption();
            }
        }

        public string Export(string desPath)
        {
            var srcDoc = new XWPFDocument(File.OpenRead(_templatePath));
            var desDoc = new XWPFDocument(File.OpenRead(desPath));
            for (var srcIndex = 0; srcIndex < desDoc.Paragraphs.Count; srcIndex++)
            {
                desDoc.RemoveBodyElement(desDoc.GetPosOfParagraph(desDoc.Paragraphs[0]));
            }

            Export<T>(srcDoc, desDoc, 0, _data);

            return desPath;

        }

        public string Export<Model>(XWPFDocument srcDoc, XWPFDocument desDoc, int srcIndex,T t)
        {

            for (srcIndex = 0; srcIndex < srcDoc.Paragraphs.Count; srcIndex++)
            {
                var srcPara = srcDoc.Paragraphs[srcIndex];
                var matches = Regex.Matches(srcPara.ParagraphText, _option.Pattern);

                if (!matches.Any())
                {
                    var desPara = new XWPFParagraph(srcPara.GetCTP(), desDoc.CreateParagraph().Body);
                    desDoc.SetParagraph(desPara, _desIndex++);
                    continue;
                }

                foreach (Match match in matches)
                {
                    var filedName = match.Groups[1].Value;
                    var propertyInfo = typeof(T).GetProperties().FirstOrDefault(m => m.Name == filedName);
                    
                    if (match.Value.StartsWith(_option.StartPattern + _option.StartTag))
                    {
                        Export<propertyInfo.PropertyType>(srcDoc, desDoc, srcIndex, propertyInfo.GetValue(filedName));
                    }
                    if (match.Value.EndsWith(_option.EndTag + _option.EndPattern))
                    {
                        Export(srcDoc, desDoc, srcIndex);
                    }
                }

            }
            return string.Empty;
        }
    }
    
}
