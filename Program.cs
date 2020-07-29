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
            List<string> headerCaptions = new List<string>();
            List<string> headerValues = new List<string>();
            List<string> tableHeaders = new List<string>();
            List<List<string>> tableRows = new List<List<string>>();
            string[,] rowsArray = new string[,] { { "" }, { "" } };
            if (root.HasChildNodes)
            {

                foreach (XmlNode node in headers)
                {
                    Log(node.Attributes["Caption"].Value);
                    headerCaptions.Add(node.Attributes["Caption"].Value);
                    headerValues.Add(node.Attributes["Value"].Value);
                }
                // seting table headers from the first row
                foreach (XmlNode node in rows.FirstChild)
                {
                    tableHeaders.Add(node.Attributes["Caption"].Value);
                }
                // seting table rows as list of lists
                foreach (XmlNode rowNode in rows)
                {
                    List<string> oneRow = new List<string>();
                    foreach (XmlNode oneRowNode in rowNode)
                    {
                        oneRow.Add(oneRowNode.Attributes["Value"].Value);
                    }
                    tableRows.Add(oneRow);



                }


                Log(tableRows.ToString());


            }




            // Log(xmlData.InnerXml);



        }
        private static void Log(string text)
        {
            System.Diagnostics.Debug.WriteLine(text);
        }

    }
}
