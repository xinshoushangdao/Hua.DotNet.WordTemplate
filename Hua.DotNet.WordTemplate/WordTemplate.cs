using NPOI.SS.Formula.Functions;
using NPOI.XWPF.UserModel;
using System.Reflection;
using System.Text.RegularExpressions;
using Hua.DotNet.WordTemplate.Model;
using Match = System.Text.RegularExpressions.Match;
using System.Text;

namespace Hua.DotNet.WordTemplate
{
    /// <summary>
    /// Word template 
    /// </summary>
    public class WordTemplate<T>
    {
        private readonly string _templatePath;
        private readonly byte[] _srcFs;
        private readonly T _data;
        private readonly TemplateOption _option;
        private int _desIndex = 0;
        
        public WordTemplate(T data, string templatePath, TemplateOption? option = null)
        {
            this._templatePath = templatePath;
            _srcFs = Encoding.Default.GetBytes(File.ReadAllText(_templatePath));
            this._data = data;
            if (option == null)
            {
                this._option = new TemplateOption();
            }
        }

        public string Export(string desPath)
        {
            //var srcDoc = new XWPFDocument(File.OpenRead(_templatePath));
            var desDoc = new XWPFDocument(File.OpenRead(_templatePath));
            RemoveParas(desDoc);
            ExportModel(_data, desDoc, 0);
            desDoc.Write(File.Open(desPath,FileMode.CreateNew));
            return desPath;
        }

        private void RemoveParas(XWPFDocument desDoc)
        {
            var count = desDoc.Paragraphs.Count;
            for (var srcIndex = 0; srcIndex < count; srcIndex++)
            {
                desDoc.RemoveBodyElement(desDoc.GetPosOfParagraph(desDoc.Paragraphs[0]));
            }
        }

        public int ExportModel<TData>(TData model,XWPFDocument desDoc, int srcIndex = 0)
        {
            using var srcDoc = new XWPFDocument(File.OpenRead(_templatePath));

            var isStartLine = typeof(TData) != typeof(T);
            var offset = 0;
            var tData = typeof(TData);
            for (; srcIndex+ offset < srcDoc.Paragraphs.Count; offset++)
            {
                
                var srcPara = srcDoc.Paragraphs[srcIndex + offset];
                var desPara = new XWPFParagraph(srcDoc.Paragraphs[srcIndex + offset].GetCTP(), desDoc.CreateParagraph().Body);

                //Demo: ${Name} ${Start:Data1} ${End:Data1}
                var matches = Regex.Matches(srcPara.ParagraphText, _option.Pattern);

                //0.No Tag
                if (!matches.Any())
                {
                    desDoc.SetParagraph(desPara, _desIndex++);
                    continue;
                }
                
                //1.deal Tags
                foreach (Match match in matches)
                {
                    string filedName;
                    PropertyInfo propertyInfo;

                    //递归属性时的出口
                    if (match.Value.StartsWith(_option.StartPattern + _option.EndTag))
                    {
                        desPara.ReplaceText(match.Value,string.Empty);
                        foreach (var match1 in matches.Where(m=>m.Value!=match.Value))
                        {
                            desPara.ReplaceText(match1.Value, string.Empty);
                        }
                        desDoc.SetParagraph(desPara, _desIndex++);
                        return offset;
                    }

                    //有开始标志，准备递归
                    if (match.Value.StartsWith(_option.StartPattern + _option.StartTag))
                    {
                        if (isStartLine)
                        {
                            desPara.ReplaceText(match.Value, string.Empty);
                            continue;
                        }
                        #region 获取属性
                        var index = match.Value.IndexOf(_option.Separator, StringComparison.Ordinal)+1;
                        filedName = match.Value
                            .Substring(index,
                                match.Value.Length - index - _option.EndPattern.Length);
                        propertyInfo = tData.GetProperties().FirstOrDefault(m => m.Name == filedName);
                        if (!typeof(IEnumerable<object>).IsAssignableFrom(propertyInfo.PropertyType))
                        {
                            throw new Exception("迭代类型必须为IEnumerable<>");
                        }

                        var value = propertyInfo.GetValue(model);
                        if (value==null) continue;
                        #endregion

                        var list = value as IEnumerable<object>;
                        var methodType = typeof(WordTemplate<T>).GetMethod("ExportModel")!
                            .MakeGenericMethod(propertyInfo.PropertyType.GetGenericArguments().First());
                        var coff = 0;
                        foreach (var obj in list)
                        {
                            //递归入口
                            coff = (int)methodType.Invoke(this, new [] { obj, desDoc, srcIndex + offset })!;
                        }
                        offset += coff;
                        goto nextLine;
                    }

                    //没有开始/结束标志，赋值属性
                    filedName = match.Value
                        .Substring(_option.StartPattern.Length,
                            match.Value.Length - _option.StartPattern.Length - _option.EndPattern.Length);
                    propertyInfo = tData.GetProperties().FirstOrDefault(m => m.Name == filedName);
                    if (propertyInfo == null) continue;
                    desPara.ReplaceText(match.Value, propertyInfo.GetValue(model)!.ToString());
                }
                desDoc.SetParagraph(desPara, _desIndex++); 
                nextLine:;
            }

            return 0;
        }
    }
}