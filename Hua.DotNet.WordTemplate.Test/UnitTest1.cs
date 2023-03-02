using System.Text.RegularExpressions;

namespace Hua.DotNet.WordTemplate.Test
{
    public class UnitTest1
    {
        [Fact]
        public void Test1()
        {
            var template = new WordTemplate<Data>(new Data(),"");
        }
        [Fact]
        public void Test2()
        {
            string input = "����һ�� {b:labla} ʾ����";
            string pattern = @"\{[^\}]*\}";

            MatchCollection matches = Regex.Matches(input, pattern);
            foreach (Match match in matches)
            {
                Console.WriteLine(match.Value);
            }

        }
}
}