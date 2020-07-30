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
            Random rn = new Random();
            int rand = rn.Next(1, 100);
            string fileName = "יצוא בדיקות מעבדה" + " A " + rand + ".xlsx";

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




            //Template path from library : $"{AppDomain.CurrentDomain.BaseDirectory}GraphTemplate.xlsx"
            using (ExcelPackage p = new ExcelPackage(template, true))
            {

                //default for sheets is 1 for dotnetcore
                ExcelWorksheet wsBase = p.Workbook.Worksheets[1];
                ExcelWorksheet ws = p.Workbook.Worksheets[2];
                int totalRows = rowsData.ChildNodes.Count + 1;
                int totalCols = rowsData.FirstChild.ChildNodes.Count;

                for (int rownumber = 0; rownumber < totalRows; rownumber++)
                {
                    ws.InsertRow(2 + rownumber, 1);
                    ws.Cells[4 + rownumber, 1, 4 + rownumber, totalCols].Copy(ws.Cells[2 + rownumber, 1, 2 + rownumber, totalCols]);
                }
                for (int colnumber = 0; colnumber < totalCols; colnumber++)
                {
                    ws.InsertColumn(2 + colnumber, 1);
                    ws.Cells[1, 3 + colnumber, totalRows, 3 + colnumber].Copy(ws.Cells[1, 2 + colnumber, totalRows, 2 + colnumber]);
                }
                ws.Cells[totalRows + 1, 1, totalRows + 12, totalCols + 20].Clear();
                ws.Cells[1, totalCols + 1, totalRows + 10, totalCols + 10].Clear();
                /* ws.InsertRow(2, rowsData.FirstChild.ChildNodes.Count);
                 ws.InsertColumn(2, rowsData.ChildNodes.Count);*/
                // first row headers
                // seting table captions from the first row data

                int ColIndex = 1; //First cell number is 1
                foreach (XmlNode nodeTableCaption in rowsData.FirstChild)
                {
                    ws.Cells[1, ColIndex].Value = nodeTableCaption.Attributes["Caption"].Value;
                    ColIndex += 1;
                }


                int RowIndex = 2; //Start rows after caption row

                foreach (XmlNode rowNode in rowsData)
                {
                    ColIndex = 1; //First cell number is 1
                    foreach (XmlNode oneRowNode in rowNode)
                    {
                        var value = oneRowNode.Attributes["Value"].Value;
                        decimal decimalValue;

                        if (decimal.TryParse(value, out decimalValue))
                        {
                            ws.Cells[RowIndex, ColIndex].Value = decimalValue;
                        }
                        else
                        {
                            ws.Cells[RowIndex, ColIndex].Value = value;
                        };

                        ColIndex += 1;
                    }
                    //tableRows.Add(oneRow);

                    RowIndex += 1;

                }
                // changing headers in cells

                foreach (var worksheetCell in wsBase.Cells)
                {
                    if (worksheetCell?.Value?.ToString() == "<TABLE>")
                    {
                        worksheetCell.Copy(ws.Cells[1, 1, totalRows + 1, totalCols + 1])
                    }

                    foreach (XmlNode nodeHeaderField in headerData)
                    {
                        if (worksheetCell?.Value?.ToString() == "<" + nodeHeaderField.Attributes["Caption"].Value + ">")
                        {
                            worksheetCell.Value = nodeHeaderField.Attributes["Value"].Value;
                        }

                    }




                }
                ws.Cells[ws.Dimension.Address].AutoFitColumns();

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
