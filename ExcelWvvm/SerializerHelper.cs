using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;
using System.IO;
using System.Xml;

namespace ExcelWvvm
{
    public class SerializerHelper
    {
        public static string ToXml(object graph)
        {
            DataContractSerializer dcs = new DataContractSerializer(graph.GetType());
            XmlWriterSettings xs = new XmlWriterSettings();
            xs.Indent = true;
            xs.OmitXmlDeclaration = true;
            xs.CheckCharacters = true;
            StringBuilder sb = new StringBuilder();
            using(XmlWriter writer = XmlWriter.Create(sb, xs))
            {
                dcs.WriteObject(writer, graph);
            }
            return sb.ToString();
        }

        //Be careful DataContractSerializer is order sensitive, see http://stackoverflow.com/questions/1727682/wcf-disable-deserialization-order-sensitivity
        //We may implement from XmlObjectSerializer to fix the order issue.
        public static T FromXml<T>(object graph, string xml)
        {
            DataContractSerializer dcs = new DataContractSerializer(graph.GetType());
            StringReader sr = new StringReader(xml);
            XmlTextReader reader = new XmlTextReader(sr);
            return (T)dcs.ReadObject(reader);
        }
    }
}
