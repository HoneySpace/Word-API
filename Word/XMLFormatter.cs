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

        public void GetDoc()
        {
            Console.WriteLine(WordXML.W + "body");
            XElement body = doc.Root.Elements(WordXML.Pkg + "part")
                .FirstOrDefault(n => n.Attribute(WordXML.Pkg + "name").Value == "/word/document.xml")
                .Element(WordXML.Pkg + "xmlData")
                .Element(WordXML.W + "document")
                .Element(WordXML.W + "body");
            foreach (XElement p in body.Elements(WordXML.W + "p"))
            {
                p.Elements(WordXML.W + "proofErr").Remove();
                p.Elements(WordXML.W + "bookmarkStart").Remove();
                p.Elements(WordXML.W + "bookmarkEnd").Remove();
            }
            foreach (XElement p in body.Elements(WordXML.W + "p"))
            {
                foreach (XNode node in p.Elements(WordXML.W + "r"))
                {
                    if (CheckNode(node)) InsertData(node);
                }
            }
            doc.Save("D:/Отчёт.doc");
            Console.WriteLine("End");
            Console.ReadKey();
        }

        public bool CheckNode(XNode node)
        {
            bool open = false;
            bool close = false;
            try
            {
                open = ((XElement) node.PreviousNode).Element(WordXML.W + "t").Value.Contains(openTag);
            }
            catch (Exception) { }
            try
            {
                close = ((XElement)node.NextNode).Element(WordXML.W + "t").Value.Contains(closeTag);
            }
            catch (Exception) { }

            return open && close;

        }
        public void InsertData(XNode node)
        {
            XElement prevText = ((XElement) node.PreviousNode).Element(WordXML.W + "t");
            XElement text = ((XElement) node).Element(WordXML.W + "t");
            XElement nextText = ((XElement) node.NextNode).Element(WordXML.W + "t");

            prevText.SetAttributeValue(WordXML.Xml + "space", "preserve");
            nextText.SetAttributeValue(WordXML.Xml + "space", "preserve");
            prevText.Value = prevText.Value.Replace("{{", "");
            nextText.Value = nextText.Value.Replace("}}", "");

            text.Value = "ВСТАВИЛ";

        }
    }
}
