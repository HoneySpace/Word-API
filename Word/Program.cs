using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using word = Microsoft.Office.Interop.Word;

namespace Word
{
    class Program
    {
        static word.Application wordapp = new word.Application();

        public static void Close()
        {
            Object saveChanges = word.WdSaveOptions.wdPromptToSaveChanges;
            Object originalFormat = word.WdOriginalFormat.wdWordDocument;
            Object routeDocument = Type.Missing;
            wordapp.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
        }



        static void Main(string[] args)
        {
            word.Documents worddocuments;
            word.Document worddocument;
            wordapp.Visible = true;
            string template = @"D:\a2.doc";

            Object newTemplate = false;
            Object documentType = word.WdNewDocumentType.wdNewBlankDocument;
            Object visible = true;
            wordapp.Documents.Add(template, newTemplate, ref documentType, ref visible);

            worddocuments = wordapp.Documents;
            worddocument = worddocuments.get_Item(1);
            worddocument.Activate();


            wordapp.Selection.HomeKey(word.WdUnits.wdStory, word.WdMovementType.wdMove);
            bool mb = true;
            string ObjectTemplate = "";
            Dictionary<string, string> variables = new Dictionary<string, string>();

            while (mb)
            {
                wordapp.Selection.MoveRight(word.WdUnits.wdWord, 1, word.WdMovementType.wdExtend);
                string gg = wordapp.Selection.Text.Trim();
                if (gg == "]")
                {
                    wordapp.Selection.MoveRight(word.WdUnits.wdCharacter, 1, word.WdMovementType.wdMove);
                    wordapp.Selection.MoveLeft(word.WdUnits.wdWord, 3, word.WdMovementType.wdExtend);
                    string text = wordapp.Selection.Text.Trim();
                    if (Istag(text))
                    {
                        if (text == "[End]")
                        {
                            wordapp.Selection.MoveLeft(word.WdUnits.wdCharacter, 1, word.WdMovementType.wdExtend);
                            wordapp.Selection.Text = "";
                            mb = false;
                        }
                        if (text == "[Name]")
                        {
                            wordapp.Selection.Text = "Вася Пупкин";
                            wordapp.Selection.Font.Size = 14;
                            //wordapp.Selection.Font.Color = word.WdColor.wdColorBlue;
                            //wordapp.Selection.Font.Size = 20;
                            //wordapp.Selection.Font.Name = "Arial";
                            //wordapp.Selection.Font.Italic = 1;
                            //wordapp.Selection.Font.Bold = 0;
                            //wordapp.Selection.Font.Underline = word.WdUnderline.wdUnderlineSingle;
                            //wordapp.Selection.Font.UnderlineColor = Word.WdColor.wdColorDarkRed;
                        }
                        if (text == $"[{Tag.Template}]")
                        {
                            wordapp.Selection.Text = ObjectTemplate;
                            wordapp.Selection.HomeKey(word.WdUnits.wdStory, word.WdMovementType.wdMove);
                        }
                        if (text == $"[{Tag.Age}]")
                        {
                            wordapp.Selection.Text = 20.ToString();
                        }
                        if (text == $"[{Tag.Object}]")
                        {
                            ObjectTemplate = InnerText(text);
                        }
                        if (text == $"[{Tag.Set}]")
                        {
                            string nextTag = FindTag();
                            string name = "";
                            string value = "";

                            while (nextTag != $"[{Tag.Set}]")
                            {
                                if (nextTag == $"[{Tag.Name}]")
                                {
                                    name = InnerText(nextTag);
                                }
                                if (nextTag == $"[{Tag.Value}]")
                                {
                                    value = InnerText(nextTag);
                                }
                                nextTag = FindTag();
                            }

                            string i = variables.Keys.ToList().Find(key => key == name);
                            if (i == null)
                                variables.Add(name, value);
                            else
                                variables[i] = value;

                            SelecetTagBack();
                            wordapp.Selection.Text = "";
                        }
                        if (text == $"[{Tag.Variable}]")
                        {
                            string nextTag = FindTag();
                            string name = "";

                            while (nextTag != $"[{Tag.Variable}]")
                            {
                                if (nextTag == $"[{Tag.Name}]")
                                {
                                    name = InnerText(nextTag);
                                }
                                nextTag = FindTag();
                            }

                            //wordapp.Selection.MoveRight(word.WdUnits.wdCharacter, 1, word.WdMovementType.wdMove);
                            SelecetTagBack();
                            wordapp.Selection.Text = "";

                            wordapp.Selection.InsertAfter(variables[name]);
                        }
                    }
                }
                wordapp.Selection.MoveRight(word.WdUnits.wdCharacter, 1, word.WdMovementType.wdMove);
                Thread.Sleep(100);
            }

            //int i = 2;
            //var wordparagraphs = worddocument.Paragraphs;
            //var wordparagraph = wordparagraphs[i];
            //wordparagraph.Range.Text = $"Текст который мы выводим в {i} абзац";
            string outputPath = @"D:\Отчёт.doc";

            try
            {
                worddocument.SaveAs(outputPath, word.WdSaveFormat.wdFormatDocument);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }


            Close();

            Console.WriteLine("Завершено");

            Console.ReadLine();
        }

        static string FindTag()
        {
            bool InLoop = false;
            int loopDepth = 20;
            int loopCount = 0;
            string prev = "";
            wordapp.Selection.MoveRight(word.WdUnits.wdCharacter, 0, word.WdMovementType.wdExtend);

            while (!InLoop)
            {
                wordapp.Selection.MoveRight(word.WdUnits.wdWord, 1, word.WdMovementType.wdExtend);
                string gg = wordapp.Selection.Text.Trim();
                if (gg == "]")
                {
                    wordapp.Selection.MoveRight(word.WdUnits.wdCharacter, 1, word.WdMovementType.wdMove);
                    wordapp.Selection.MoveLeft(word.WdUnits.wdWord, 3, word.WdMovementType.wdExtend);
                    string text = wordapp.Selection.Text.Trim();
                    if (Istag(text))
                    {
                        return text;
                    }
                }
                if (gg == prev) loopCount++;
                prev = gg;
                wordapp.Selection.MoveRight(word.WdUnits.wdCharacter, 1, word.WdMovementType.wdMove);

                if (loopCount == loopDepth)
                {
                    InLoop = true;
                    throw new Exception("Привыешено кол-во повторений, обнаружен бесконечный цикл. В тексте не обнаружен тэг.");
                }

            }

            return "";
        }

        static void SelecetTagBack()
        {
            wordapp.Selection.MoveRight(word.WdUnits.wdCharacter, 1, word.WdMovementType.wdMove);

            bool InLoop = false;
            int loopDepth = 20;
            int loopCount = 0;
            string prev = "";
            string newPart = "";
            StringBuilder builder = new StringBuilder();
            string tag = null;
            string openTag = null;

            bool isStarted = false;

            while (!InLoop)
            {
                wordapp.Selection.MoveLeft(word.WdUnits.wdWord, 1, word.WdMovementType.wdExtend);

                if (prev != "")
                    newPart = wordapp.Selection.Text.Replace(prev, "");
                prev = wordapp.Selection.Text;
                newPart = newPart.Trim();

                if (isStarted)
                {
                    builder.Insert(0, newPart);
                }

                if (newPart == "]")
                {
                    isStarted = true;
                    builder.Insert(0, newPart);
                }

                if (prev == "]")
                {
                    isStarted = true;
                    builder.Insert(0, prev);
                }

                if (newPart == "[")
                {
                    isStarted = false;



                    string text = builder.ToString();
                    if (Istag(text))
                    {
                        if (builder.ToString() == openTag)
                        {
                            InLoop = true;
                        }

                        if (openTag == null)
                            openTag = builder.ToString();

                        builder.Clear();
                    }




                }
            }
        }

        static string InnerText(string closingTag)
        {
            string prev = "";
            string newPart = "";
            string tag = "";
            bool isClosed = false;
            string output;
            wordapp.Selection.Text = "";
            wordapp.Selection.MoveRight(word.WdUnits.wdWord, 1, word.WdMovementType.wdExtend);
            wordapp.Selection.Text = wordapp.Selection.Text == "\r" ? "" : wordapp.Selection.Text;

            bool isStarted = false;

            while (!isClosed)
            {
                if (prev != "")
                    newPart = wordapp.Selection.Text.Replace(prev, "");
                prev = wordapp.Selection.Text;

                if (newPart == "[")
                    isStarted = true;

                if (isStarted)
                    tag += newPart;

                if (newPart.Trim() == "]")
                {
                    if (Istag(tag.Trim()))
                    {
                        if (tag.Trim() == closingTag)
                        {
                            output = prev.Replace(tag, "");
                            wordapp.Selection.Text = "";
                            isClosed = true;
                            return output;
                        }
                    }
                    isStarted = false;
                    tag = "";
                }

                if (!isClosed)
                    wordapp.Selection.MoveRight(word.WdUnits.wdWord, 1, word.WdMovementType.wdExtend);
            }
            return "";
        }

        static bool Istag(string value)
        {
            foreach (Tag tag in Enum.GetValues(typeof(Tag)))
            {
                if (value == $"[{tag}]")
                {
                    return true;
                }
            }
            return false;
        }
    }

    public enum Tag
    {
        Name,
        End,
        Object,
        Template,
        Age,
        Set,
        Variable,
        Value,
    }
}
