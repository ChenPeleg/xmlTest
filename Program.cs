using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;




namespace ConsoleApp1

{
    class Program
    {

        static void Main(string[] args)
        {

            XmlDocument xmlData = new XmlDocument();
            string path = $"{AppDomain.CurrentDomain.BaseDirectory}";
            string xmlPath = Path.GetFullPath(Path.Combine(path, @"..\..\data.xml"));
            xmlData.Load(xmlPath);
            XmlNode root = xmlData.FirstChild;

            Log("----------\n\n");
            Log("Running");

            XmlNode headers = root.SelectSingleNode("Header");
            XmlNode rows = root.SelectSingleNode("Rows");
            string[,] headersArray = new string[,] { { "" }, { "" } };
            string[,] rowsArray = new string[,] { { "" }, { "" } };
            if (root.HasChildNodes)
            {
                // Log(headers.InnerXml);
                for (int i = 0; i < headers.ChildNodes.Count; i++)
                {

                }

                Log(rows.InnerXml);

            }




            // Log(xmlData.InnerXml);



        }
        private static void Log(string text)
        {
            System.Diagnostics.Debug.WriteLine(text);
        }

    }
}
