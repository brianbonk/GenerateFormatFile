using System.Linq;
using System.IO;
using System.Data;
using System.Xml.Linq;
using System;

namespace Common
{
    public class HandleFile
    {
        public static XDocument GenerateXMLFile(string path, string filename, string delimiter)
        {
            /*
             *Only the t, n, r, 0 and '\0' characters work with the backslash escape character to produce a control character.
             */
            var splitByDelimiter = delimiter;
            splitByDelimiter = splitByDelimiter.Replace("\\t", "\t");
            splitByDelimiter = splitByDelimiter.Replace("\\n", "\n");
            splitByDelimiter = splitByDelimiter.Replace("\\r", "\r");
            splitByDelimiter = splitByDelimiter.Replace("\\0", "\0");

            var reader = new StreamReader(File.OpenRead(path + filename));
            string[] values = reader.ReadLine().Replace(@"""", @"").Split(splitByDelimiter.ToCharArray(), StringSplitOptions.None);

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
                                    new XAttribute("MAX_LENGTH", "8000"),
                                    new XAttribute("COLLATION", "Latin1_General_CI_AS")
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