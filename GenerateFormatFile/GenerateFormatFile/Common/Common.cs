using System.Linq;
using System.IO;
using System.Data;
using System.Xml.Linq;

namespace Common
{
    public class HandleFile
    {
        public static XDocument GenerateXMLFile(string path, string filename, string delimiter)
        {
            var reader = new StreamReader(File.OpenRead(path + filename));
            string[] values = reader.ReadLine().Replace(@"""", @"").Split(';');

            var ns = XNamespace.Get("http://schemas.microsoft.com/sqlserver/2004/bulkload/format");
            var nsi = XNamespace.Get("http://www.w3.org/2001/XMLSchema-instance");

            XDocument doc =
                  new XDocument(
                    new XElement(ns + "BCPFORMAT",
                        new XAttribute(XNamespace.Xmlns + "xsi", nsi),
                            new XElement("RECORD",
                            values.Select((v, index) =>
                                new XElement("FIELD",
                                    new XAttribute("ID", index + 1),
                                    new XAttribute(nsi + "type", "CharTerm"),
                                    new XAttribute("TERMINATOR", index == values.Length - 1 ? "\\r\\n" : delimiter),
                                    new XAttribute("MAX_LENGTH", "510"),
                                    new XAttribute("COLLATION", "Latin1_General_CI_AS_KS_WS")
                                    )
                                )
                            )
                        , new XElement("ROW",
                            values.Select((v, index) =>
                                new XElement("COLUMN",
                                    new XAttribute("SOURCE", index + 1),
                                    new XAttribute("NAME", v),
                                    new XAttribute(nsi + "type", "SQLNVARCHAR")
                                )
                            )
                        )
                    )
                );
            return doc;
        }
    }
}