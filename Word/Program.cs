﻿using System;
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
            //COMFormatter.Apply();
            XMLFormatter xmlFormatter = new XMLFormatter();
            xmlFormatter.GetDoc();
        }
    }
}
