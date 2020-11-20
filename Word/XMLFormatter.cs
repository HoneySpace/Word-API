using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Word
{
    class XMLFormatter
    {
        private XDocument doc = XDocument.Load(@"D:/a2.xml");
        private string openTag = "{{";
        private string closeTag = "}}";

        public void GetDoc()
        {

            XElement body = doc.Element("w:body");
            Console.WriteLine(body);

            Console.ReadKey();
        }
    }
}
