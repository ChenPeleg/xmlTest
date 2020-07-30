using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace xmlToXls

{
    class Program
    {

        static void Main(string[] args)
        {

            XmlDocument sampleXmlData = new XmlDocument();
            string path = $"{AppDomain.CurrentDomain.BaseDirectory}";
            string xmlPath = Path.GetFullPath(Path.Combine(path, @"..\..\data.xml"));
            sampleXmlData.Load(xmlPath);
            XmlNode xmlData = sampleXmlData.FirstChild;
            string templatePath = Path.GetFullPath(Path.Combine(path, @"..\..\GraphTemplate.xlsx"));
            FileInfo template = new FileInfo(templatePath);
            createXlsx(xmlData, template);


        }


        private static void createXlsx(XmlNode xmlData, FileInfo template)
        {
            XmlNode headers = xmlData.SelectSingleNode("Header");
            XmlNode rows = xmlData.SelectSingleNode("Rows");
            List<string> headerCaptions = new List<string>();
            List<string> headerValues = new List<string>();
            List<string> tableHeaders = new List<string>();
            List<List<string>> tableRows = new List<List<string>>();
            string[,] rowsArray = new string[,] { { "" }, { "" } };
            if (xmlData.HasChildNodes)
            {

                foreach (XmlNode node in headers)
                {

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

            }

            string path = $"{AppDomain.CurrentDomain.BaseDirectory}";
            string outpudirPath = Path.GetFullPath(Path.Combine(path, @"..\..\SampleOutPut"));
            Utils.OutputDir = new DirectoryInfo(outpudirPath);

            //$"{AppDomain.CurrentDomain.BaseDirectory}SampleApp"



            string fileName = String.Join("- ", headerValues);
            Random rnd = new Random();
            int rand = rnd.Next(1, 10000);
            fileName = fileName + rand + ".xlsx";
            fileName = fileName.Replace("/", "-").Replace("\\", "-").
                Replace(@"\", "-").Replace("//", "-");

            //Template path from library : $"{AppDomain.CurrentDomain.BaseDirectory}GraphTemplate.xlsx"
            using (ExcelPackage p = new ExcelPackage(template, true))
            {
                //Set up the headers
                //default for sheets is 1 for dotnetcore
                ExcelWorksheet ws = p.Workbook.Worksheets[1];
                // first row headers
                for (int i = 0; i < tableHeaders.Count; i++)
                {
                    int cell = i + 1;
                    int firstRow = 1;
                    ws.Cells[firstRow, cell].Value = tableHeaders[i];
                }
                for (int r = 0; r < tableRows.Count; r++)
                {
                    int row = r + 2;
                    List<string> thisRow = tableRows[r];
                    for (int c = 0; c < thisRow.Count; c++)
                    {
                        int cell = c + 1;
                        var value = thisRow[c];
                        decimal decimalValue;

                        if (decimal.TryParse(value, out decimalValue))
                        {
                            ws.Cells[row, cell].Value = decimalValue;
                        }
                        else
                        {
                            ws.Cells[row, cell].Value = value;
                        };



                    }

                }




                Byte[] bin = p.GetAsByteArray(); //Get the documet as a byte array from the stream and save it to disk.  (This is useful in a webapplication) ... 

                FileInfo file = Utils.GetFileInfo(fileName);  //return file.FullName;
                File.WriteAllBytes(file.FullName, bin);

                Process.Start(Utils.OutputDir.FullName);
            }

        }
        private static void Log(string text)
        {
            System.Diagnostics.Debug.WriteLine(text);
        }

    }
}
