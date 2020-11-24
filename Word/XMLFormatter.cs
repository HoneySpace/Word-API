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
        private XElement body;
        private string openTag = "{{";
        private string closeTag = "}}";
        private Random rand = new Random();

        public XMLFormatter()
        {
            body = doc.Root.Elements(WordXML.Pkg + "part")
                .FirstOrDefault(n => n.Attribute(WordXML.Pkg + "name").Value == "/word/document.xml")
                .Element(WordXML.Pkg + "xmlData")
                .Element(WordXML.W + "document")
                .Element(WordXML.W + "body");
        }

        private enum tags
        { number, animalType, action, velocity, distanceMeasure, timeMeasure }

        public void GetDoc()
        {
            Console.WriteLine(WordXML.W + "body");
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

            AppendP("Я это написал в шарпе");
            AppendP("Ещё пишу");

            doc.Save("D:/Отчёт.doc");
            Console.WriteLine("End");
            Console.ReadKey();
        }


        public void AppendP(string message)
        {
            body.Add(new XElement(WordXML.W + "p",
                body.Element(WordXML.W + "p").Attributes(),
                new XElement(WordXML.W + "r",
                    new XElement(WordXML.W + "rPr",
                        body.Element(WordXML.W + "p").Element(WordXML.W + "r").Element(WordXML.W + "rPr").Elements()),
                    new XElement(WordXML.W + "t",
                        message,
                        new XAttribute(WordXML.Xml + "space", "preserve")
                    )
                )
            ));
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


            switch (text.Value)
            {
                case "number":
                    text.Value = rand.Next(10000).ToString();
                    break;
                case "animalType":
                    text.Value = new string[4]{"рыбы", "птицы", "насекомые", "одноклеточные"}[rand.Next(4)];
                    break;
                case "action":
                    text.Value = new string[4] { "бегать", "плавать", "прыгать", "летать" }[rand.Next(4)];
                    break;
                case "velocity":
                    text.Value = rand.Next(10000).ToString();
                    break;
                case "distanceMeasure":
                    text.Value = new string[4] { "парсек", "метров", "футов", "слонов" }[rand.Next(4)];
                    break;
                case "timeMeasure":
                    text.Value = new string[4] { "милисекунду", "секунду", "час", "вечность" }[rand.Next(4)];
                    break;
                default:
                    text.Value = "undefined";
                    break;
            }
        }
    }
}
