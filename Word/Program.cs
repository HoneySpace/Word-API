using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Word
{
    class Program
    {
        static void Main(string[] args)
        {
            Random random = new Random();
            //COMFormatter.Apply();
            //XMLFormatter xmlFormatter = new XMLFormatter();
            COMFormatter comFormatter = new COMFormatter(@"D:\Temp.doc");

            //comFormatter.Apply();

            comFormatter.Replace("{ЖИВОТНОЕ}", (new string[3] {"жирафы", "слоны", "сычы" })[random.Next(0,3)]);
            comFormatter.Replace("{ДЕЙСТВИЕ}", (new string[3] {"выпивать", "драться", "кувыркаться" })[random.Next(0,3)]);
            comFormatter.Replace("{КОЛ-ВО РАЗ}", random.Next(0,20000).ToString());
            comFormatter.Replace("{РЕЗЛЬТАТ ДЕЙСТВИЯ}", (new string[3] {"получены", "рассчитаны", "придуманы" })[random.Next(0,3)]);
            comFormatter.Replace("{ТИП ДАННЫХ}", (new string[3] {"данные", "выводы", "боги" })[random.Next(0,3)]);
            
            comFormatter.ReplaceBookmark("mark", "Этот текст был закладкой");

            comFormatter.Close();
            //xmlFormatter.GetDoc();
        }
    }
}
