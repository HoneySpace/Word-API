using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;



namespace Word
{
    class XMLFormatter
    {
        private XDocument doc = XDocument.Load(@"D:/a2.xml");
        private string openTag = "{{";
        private string closeTag = "}}";

        public XMLFormatter()
        {
        }

        public void GetDoc()
        {
            Console.WriteLine(WordXML.W + "body");
            XElement body = doc.Element(WordXML.W + "body");

            Console.WriteLine(doc);
            Console.WriteLine(body);


            Console.WriteLine("End");
            Console.ReadKey();
        }
    }
}
