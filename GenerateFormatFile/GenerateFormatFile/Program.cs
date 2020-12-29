using System;
using System.IO;
using NDesk.Options;
using FormatFile.Common;
using System.Linq;

namespace GenerateFormatFile
{
    class Program
    {
        static void Main(string[] args)
        {

            string file = "";
            string delimiter = "";
            bool normalize = false;
            bool showHelp = false;
            string servername = GenerateFormatFile.Properties.Settings.Default.servername;
            string username = GenerateFormatFile.Properties.Settings.Default.username;
            string password = GenerateFormatFile.Properties.Settings.Default.password;
            string database = GenerateFormatFile.Properties.Settings.Default.databasename;
            bool LoadToSQL = GenerateFormatFile.Properties.Settings.Default.loadToSQL;

            var p = new OptionSet()
            {
                {"f|path=", "(needed) The filename/folderpath of the file or folder to be processed", (string v)=>file=v },
                {"d|delimiter=", "(needed) Defines the delimiter for the fiels", (string v)=>delimiter=v },
                {"n|normalize=", "(optional) Normalize file if rows have different number og delimiters", (bool v)=>normalize=v },
                {"help", "Show this message and end", v=>showHelp = v != null },
            };

            try
            {
                p.Parse(args);

                if(showHelp)
                {
                    ShowHelp(p);
                    return;
                }
                //Check file and folder for existance
                if (string.IsNullOrWhiteSpace(file))
                {
                    throw new OptionException("Path or filename cannot be blank or empty", file);
                }

                //Check file and folder for existance
                if (string.IsNullOrWhiteSpace(delimiter))
                {
                    throw new OptionException("Delimiter cannot be blank or empty", delimiter);
                }

                if (File.Exists(file))
                {
                    GenerateXML.HandleFile(file, delimiter, normalize);
                }
                else if (Directory.Exists(file))
                {
                    string[] fileEntries = Directory.GetFiles(file);

                    foreach (string filename in fileEntries.Where(i => i.EndsWith(".csv") | i.EndsWith(".xlsx") | i.EndsWith(".xls")))
                    {
                        GenerateXML.HandleFile(filename, delimiter, normalize);
                    }
                }
                else
                {
                    throw new OptionException("{0}' is not a valid file or path.", file);
                }

            }
            catch (OptionException e)
            {
                Console.Write("GenerateFormatFile: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `GenerateFormatFile --help' for more information");
                Console.WriteLine("Path to the file: {0}", Path.GetDirectoryName(file));
                return;
            }

        }
        static void ShowHelp(OptionSet p)
        {
            Console.WriteLine("Use: GenerateFormatFile [OPTIONS]");
            Console.WriteLine("Generates a predefined formatfile and can load data directly to SQL Server.");
            Console.WriteLine("© Brian Bønk - 2018");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }
    }
}
