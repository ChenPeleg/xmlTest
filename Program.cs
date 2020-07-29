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
            //FileInfo dataFile =  new FileInfo();
            xmlData.Load(xmlPath);
            System.Diagnostics.Debug.WriteLine(xmlData.InnerXml);
        }
    }
}
