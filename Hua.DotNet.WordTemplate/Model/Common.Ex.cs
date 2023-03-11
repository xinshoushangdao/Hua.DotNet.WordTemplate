using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Xml.Serialization;
using Newtonsoft.Json;
using NPOI.OpenXmlFormats.Wordprocessing;

namespace Hua.DotNet.WordTemplate.Model
{
    public static class Common
    {
        public static CT_P DeepClone(this CT_P t)
        {
            var namespaces = new XmlSerializerNamespaces();
            namespaces.Add("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            using var ms = new MemoryStream(); 
            var xs = new XmlSerializer(typeof(CT_P), "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            xs.Serialize(ms, t,namespaces);
            return (CT_P)xs.Deserialize(ms);
        }
    }
}