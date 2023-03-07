namespace Hua.DotNet.WordTemplate.Model
{
    public class Data
    {
        public string Name { get; set; }
        public List<DataItem> Data1 { get; set; }
    }

    public class DataItem
    {
        public string Name { get; set; }
        public List<DataItem> Data1 { get; set; }
    }
}