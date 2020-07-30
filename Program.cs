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
            //ExcelPackage excelpackage = createXlsx(xmlData, template);
            Random rnd = new Random();
            int rand = rnd.Next(1, 10000);
            string fileName = "יצוא בדיקות מעבדה" + rand.ToString() + ".xlsx";

            createXlsx(xmlData, template, fileName);

        }
        private static void createXlsx(XmlNode xmlData, FileInfo template, String fileName)
        {
            XmlNode headerData = xmlData.SelectSingleNode("Header");
            XmlNode rowsData = xmlData.SelectSingleNode("Rows");
            List<string> headerCaptions = new List<string>();
            List<string> headerValues = new List<string>();
            List<string> tableCaptions = new List<string>();
            List<List<string>> tableRows = new List<List<string>>();

            // Convert data from XML Structure to List structure
            if (xmlData.HasChildNodes)
            {

                foreach (XmlNode nodeHeaderField in headerData)
                {

                    headerCaptions.Add(nodeHeaderField.Attributes["Caption"].Value);
                    headerValues.Add(nodeHeaderField.Attributes["Value"].Value);
                }
                // seting table captions from the first row data
                foreach (XmlNode nodeTableCaption in rowsData.FirstChild)
                {
                    tableCaptions.Add(nodeTableCaption.Attributes["Caption"].Value);
                }
                // seting table rows as list of lists
                foreach (XmlNode rowNode in rowsData)
                {
                    List<string> oneRow = new List<string>();
                    foreach (XmlNode oneRowNode in rowNode)
                    {
                        oneRow.Add(oneRowNode.Attributes["Value"].Value);
                    }
                    tableRows.Add(oneRow);



                }

            }


            //$"{AppDomain.CurrentDomain.BaseDirectory}SampleApp"
            /*  string fileName = String.Join("- ", headerValues);
              Random rnd = new Random();
              int rand = rnd.Next(1, 10000);
              fileName = fileName + rand + ".xlsx";
              fileName = fileName.Replace("/", "-").Replace("\\", "-").
                  Replace(@"\", "-").Replace("//", "-");*/

            //Template path from library : $"{AppDomain.CurrentDomain.BaseDirectory}GraphTemplate.xlsx"
            using (ExcelPackage p = new ExcelPackage(template, true))
            {
                //Set up the headers
                //default for sheets is 1 for dotnetcore
                ExcelWorksheet ws = p.Workbook.Worksheets[1];
                // first row headers
                for (int i = 0; i < tableCaptions.Count; i++)
                {
                    int cell = i + 1;
                    int firstRow = 1;
                    ws.Cells[firstRow, cell].Value = tableCaptions[i];
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
                saveFile(p, fileName);
            }


        }
        private static void saveFile(ExcelPackage excelP, string fileName)
        {

            string path = $"{AppDomain.CurrentDomain.BaseDirectory}";
            string outpudirPath = Path.GetFullPath(Path.Combine(path, @"..\..\SampleOutPut"));
            Utils.OutputDir = new DirectoryInfo(outpudirPath);
            Byte[] bin = excelP.GetAsByteArray(); //Get the documet as a byte array from the stream and save it to disk.  (This is useful in a webapplication) ... 
            FileInfo file = Utils.GetFileInfo(fileName);  //return file.FullName;
            File.WriteAllBytes(file.FullName, bin);
            Process.Start(Utils.OutputDir.FullName);

        }
        private static void Log(string text)
        {
            System.Diagnostics.Debug.WriteLine(text);
        }

    }
}
