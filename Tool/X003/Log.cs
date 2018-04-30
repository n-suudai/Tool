using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExcelDiffTool
{
    public static class Log
    {
        static StreamWriter Writer = null;

        public static void Init(string file)
        {
            Writer = new StreamWriter(file, false);
        }

        public static void Term()
        {
            Writer.Close();
            Writer = null;
        }

        public static void WriteLine(string text)
        {
            Writer.WriteLine(text);
            Console.WriteLine(text);
        }

        public static string ReadLine()
        {
            string text = Console.ReadLine();

            Writer.WriteLine(text);

            return text;
        }
    }
}
