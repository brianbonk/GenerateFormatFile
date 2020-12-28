using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ExcelDataReader;


namespace FormatFile.Common
{
    public class GenerateXML
    {
        public static bool HandleFile(string file, string delimiter)
        {
            string FileNameExtension = Path.GetExtension(file);
            string FileName = Path.GetFileNameWithoutExtension(file);
            string FullPath = Path.GetDirectoryName(file);

            Console.WriteLine("FileName: {0}", FileName);
            Console.WriteLine("File-type: {0}", FileNameExtension);
            Console.WriteLine("Full path: {0}", FullPath);

            if (FileNameExtension == ".xls" || FileNameExtension == ".xlsx")
            {
                Console.WriteLine("File is Excel-format - converting Excel to CSV");
                string csvoutput = Path.ChangeExtension(file, ".csv");
                delimiter = "|";
                ExcelFileHelper.SaveAsCsv(file, csvoutput, delimiter);
                Console.WriteLine("Deletes Excel file");
                File.Delete(file);
            }

            if (delimiter == "tab")
            {
                Console.WriteLine("Delimter set to TAB - converting TAB to pipe '|'");
                string text = "";
                using (StreamReader sr = new StreamReader(file))
                {
                    int i = 0;
                    do
                    {
                        i++;
                        string line = sr.ReadLine();
                        if (line != "")
                        {
                            line = line.Replace("\t", "|");
                            text = text + line + Environment.NewLine;
                        }
                    } while (sr.EndOfStream == false);
                }
                File.WriteAllText(FullPath + "\\" + FileName + "Converted.csv", text);
                //change delimiter and filename
                delimiter = "|";
                FileName += "Converted";
                Console.WriteLine("New filename: {0}", FileName);
            }

            Console.WriteLine("Processing CSV-file and generating XML-formatfile");
            var doc = GenerateXML.GenerateXMLFile(FullPath, "\\" + FileName + ".csv", delimiter);
            StreamWriter xmlfile = File.CreateText(FullPath + "\\formatfile_" + FileName + ".xml");
            doc.Save(xmlfile);
            xmlfile.Close();

            if (GenerateFormatFile.Properties.Settings.Default.loadToSQL)
            {
                Console.WriteLine("LoadToSQL is activated - loading file to SQL Server");
                LoadToSQL.CSVFile(
                    GenerateFormatFile.Properties.Settings.Default.servername
                    , GenerateFormatFile.Properties.Settings.Default.databasename
                    , GenerateFormatFile.Properties.Settings.Default.username
                    , GenerateFormatFile.Properties.Settings.Default.password
                    , FullPath + "\\" + FileName + ".csv");
            }

            //Console.WriteLine("Delete CSV file");
            //File.Delete(FullPath + "\\" + FileName + ".csv");
            //Console.WriteLine("Delete Format-file");
            //File.Delete(FullPath + "\\formatfile_" + FileName + ".xml");

            return true;
        }
        public static XDocument GenerateXMLFile(string path, string filename, string delimiter)
        {
            var reader = new StreamReader(File.OpenRead(path + filename));
            string[] values = reader.ReadLine().Replace(@"""", @"").Split(delimiter[0]);

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
                                    new XAttribute("COLLATION", "")
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
            reader.Close();
            return doc;
        }
    }
    public class ExcelFileHelper
    {
        public static bool SaveAsCsv(string excelFilePath, string destinationCsvFilePath, string delimiter)
        {

            using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                IExcelDataReader reader = null;
                if (excelFilePath.EndsWith(".xls"))
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (excelFilePath.EndsWith(".xlsx"))
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                if (reader == null)
                    return false;

                var ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = false
                    }
                });

                var csvContent = string.Empty;
                int row_no = 0;
                while (row_no < ds.Tables[0].Rows.Count)
                {
                    var arr = new List<string>();
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        arr.Add(Regex.Replace(ds.Tables[0].Rows[row_no][i].ToString(), @"(\n)", string.Empty));
                    }
                    row_no++;
                    csvContent += string.Join(delimiter, arr) + "\r\n";
                }
                StreamWriter csv = new StreamWriter(destinationCsvFilePath, false);
                csv.Write(csvContent);
                csv.Close();
                csv.Dispose();
                return true;
            }
        }
    }

    public class LoadToSQL
    {
        public static bool CSVFile(string servername, string database, string username, string password, string filename)
        {
            string connectionstring = @"server=" + servername + ";Database=" + database + ";";
            if (string.IsNullOrEmpty(password))
            {
                connectionstring += "Integrated Security = true";
            }
            else
            {
                connectionstring += "User=" + username + ";Password=" + password;
            }

            string query = @"exec dbo.bulkLoadCSV @file = '" + filename + "'";

            using (SqlConnection connection = new SqlConnection(connectionstring))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();
                command.ExecuteReader();
                connection.Close();
            }

            return true;
        }
    }
}