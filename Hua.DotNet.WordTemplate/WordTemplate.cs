using NPOI.XWPF.UserModel;
using System.Reflection;
using System.Text.RegularExpressions;

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
        
        public WordTemplate(T data, string templatePath, TemplateOption? option = null)
        {
            this._templatePath = templatePath;
            this._data = data;
            if (option == null)
            {
                this._option = new TemplateOption();
            }
        }

        public string Export(string desPath)
        {
            var srcDoc = new XWPFDocument(File.OpenRead(_templatePath));
            var desDoc = new XWPFDocument(File.OpenRead(_templatePath));
            for (var srcIndex = 0; srcIndex < desDoc.Paragraphs.Count; srcIndex++)
            {
                desDoc.RemoveBodyElement(desDoc.GetPosOfParagraph(desDoc.Paragraphs[0]));
            }

            ExportModel(_data,srcDoc, desDoc, 0);
            using var fs =  File.Create(desPath);
            desDoc.Write(fs);
            desDoc.Close();
            fs.Close();
            return desPath;
        }

        public void ExportModel<TData>(TData model, XWPFDocument srcDoc, XWPFDocument desDoc, int srcIndex)
        {
            for (srcIndex = 0; srcIndex < srcDoc.Paragraphs.Count; srcIndex++)
            {
                var srcPara = srcDoc.Paragraphs[srcIndex];
                var desPara = new XWPFParagraph(srcPara.GetCTP(), desDoc.CreateParagraph().Body);


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

                    var filedName = string.Empty;
                    PropertyInfo? propertyInfo = null;
                    var type = typeof(TData);

                    if (match.Value.StartsWith(_option.StartPattern + _option.StartTag))
                    {
                        var index = match.Value.IndexOf(_option.Separator, StringComparison.Ordinal)+1;
                        filedName = match.Groups[0].Value
                            .Substring(index,
                                match.Groups[0].Value.Length - index - _option.EndPattern.Length);

                        propertyInfo = type.GetProperties().FirstOrDefault(m => m.Name == filedName);
                        if (!typeof(IEnumerable<object>).IsAssignableFrom(propertyInfo.PropertyType))
                        {
                            throw new Exception("迭代类型必须为IEnumerable<>");
                        }

                        // 定义泛型委托类型
                        //递归入口
                        var value = propertyInfo.GetValue(model);
                        if (value==null) continue;
                        var list = value as IEnumerable<object>;
                        var methodType = typeof(WordTemplate<T>).GetMethod("ExportModel")!
                            .MakeGenericMethod(propertyInfo.PropertyType.GetGenericArguments().First());
                        foreach (var obj in list)
                        {
                            methodType.Invoke(this, new [] { obj, srcDoc, desDoc, srcIndex });
                        }
                        return;
                    }

                    if (match.Value.EndsWith(_option.StartPattern + _option.EndTag))
                    {
                        var index = match.Value.IndexOf(_option.Separator, StringComparison.Ordinal) + 1;
                        filedName = match.Groups[0].Value
                            .Substring(index,
                                match.Groups[0].Value.Length - index - _option.EndPattern.Length);

                        var para = new XWPFParagraph(srcPara.GetCTP(), desDoc.CreateParagraph().Body);
                        para.ReplaceText(match.Value, string.Empty);
                        desDoc.SetParagraph(para, _desIndex++);
                    }


                    filedName = match.Groups[0].Value
                        .Substring(_option.StartPattern.Length,
                            match.Groups[0].Value.Length - _option.StartPattern.Length - _option.EndPattern.Length);
                    propertyInfo = type.GetProperties().FirstOrDefault(m => m.Name == filedName);
                    if (propertyInfo != null)
                    {
                        desPara.ReplaceText(match.Groups[0].Value, propertyInfo.GetValue(model)!.ToString());
                        desDoc.SetParagraph(desPara, _desIndex++);
                    }

                }
            }
        }
    }
}