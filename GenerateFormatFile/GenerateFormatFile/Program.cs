using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Xml.Linq;
using System.Xml;
using NDesk.Options;
using Common;

namespace GenerateFormatFile
{
    class Program
    {
        static void Main(string[] args)
        {

            string path = "";
            string filename = "";
            string outputfile = "";
            string delimiter = "";
            bool showHelp = false;

            var p = new OptionSet()
            {
                {"p|path=", "(req) The path to the folder containing file", (string v)=>path=v },
                {"f|filename=", "(req) The filename to read the metadata from", (string v)=>filename=v },
                {"o|outputfile=", "(req) The output filename (the format file)", (string v)=>outputfile=v },
                {"d|delimiter=", "(req) The column delimiter for the header row", (string v)=>delimiter=v },
                {"help", "show help and close", v=>showHelp = v != null },
            };

            try
            {
                p.Parse(args);

                if(showHelp)
                {
                    ShowHelp(p);
                    return;
                }
                if (string.IsNullOrWhiteSpace(path))
                {
                    throw new OptionException("Path cannot be blank or empty.", "sti");
                }
                if (!string.IsNullOrWhiteSpace(path) && !Directory.Exists(path))
                {
                    throw new OptionException("The path '{path}' does not exist.", "sti");
                }
                if (string.IsNullOrWhiteSpace(outputfile))
                {
                    throw new OptionException("The output file '{outputfile}' cannot be blank or empty.", "outputfile");
                }
                if (string.IsNullOrWhiteSpace(filename))
                {
                    throw new OptionException("Filename cannot be blank or empty.","filename");
                }
                if(!string.IsNullOrWhiteSpace(filename) && !File.Exists(path + filename))
                {
                    throw new OptionException("The file '{filename}' does not exist.", "filename");
                }
                if (!string.IsNullOrWhiteSpace(delimiter))
                {
                    throw new OptionException("The delimiter cannot be blank or empty.", "filename");
                }

            }
            catch (OptionException e)
            {
                Console.Write("GenerateFormatFile: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `GenerateFormatFile --help' for more information");
                return;
            }

            //do magic stuff
            var doc = HandleFile.GenerateXMLFile(path, filename, delimiter);
            doc.Save(new StreamWriter(File.OpenWrite(path + outputfile)));

        }
        static void ShowHelp(OptionSet p)
        {
            Console.WriteLine("Use: GenerateFormatFile [OPTIONS]");
            Console.WriteLine("Generates a formatfile to be used in dynamic import of flatfiles.");
            Console.WriteLine("© bonk.dk 2017");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }
    }
}
