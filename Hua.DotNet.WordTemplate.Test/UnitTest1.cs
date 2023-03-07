using System.Diagnostics;
using System.Reflection;
using System.Text.RegularExpressions;
using Hua.DotNet.WordTemplate.Model;
using Microsoft.VisualStudio.TestPlatform.TestHost;
using NPOI.XWPF.UserModel;

namespace Hua.DotNet.WordTemplate.Test
{
    public class UnitTest1
    {
        delegate void MultiParameterDelegate<T1>(T1 arg1, string arg2);
        [Fact]
        public void Test1()
        {
            var template = new WordTemplate<Data>(new Data(),"");
        }
        [Fact]
        public void Test2()
        {
            string input = "这是一个 {b:labla} 示例。";
            string pattern = @"\{[^\}]*\}";

            MatchCollection matches = Regex.Matches(input, pattern);
            foreach (Match match in matches)
            {
                Console.WriteLine(match.Value);
            }

        }
        [Fact]
        public void Test3()
        {
            Type delegateType = typeof(MultiParameterDelegate<>).MakeGenericType(typeof(int));
            MethodInfo methodInfo = typeof(UnitTest1).GetMethod("HandleMultiParameters");

            var handler = Delegate.CreateDelegate(delegateType, methodInfo);
            handler.DynamicInvoke(42, "hello");
        }

        static void HandleMultiParameters(int arg1, string arg2)
        {
            Console.WriteLine($"arg1: {arg1}, arg2: {arg2}");
        }

        [Fact]
        public void TestGen()
        {
            var i = 0;
            var data = new Data()
            {
                Name = $"Name:{i++}",
                Data1 = new List<DataItem>()
            };
            for (var j = 0; j < 10; j++)
            {
                data.Data1.Add(new DataItem()
                {
                    Name = $"Name:{j++}"
                });
            }

            var templatePath = "Template.docx";
            var descPath = $"{DateTime.Now:yyyy_MM_dd_HH_mm_ss}.docx";
            var template = new WordTemplate<Data>(data, templatePath);
            template.Export(descPath);
            Console.WriteLine(descPath);
        }
    }
}