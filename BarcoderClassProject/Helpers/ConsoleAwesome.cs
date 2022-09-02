using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BarcoderClassProject.Helpers
{
    public static class ConsoleAwesome
    {
        public static void Write(string text)
        {
            Console.Write(text);
        }

        public static void WriteLine(string text)
        {
            Console.WriteLine(text);
        }

        public static void Write(string text, ConsoleColor color)
        {
            ConsoleColor prevColor = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.Write(text);
            Console.ForegroundColor = prevColor;
        }

        public static void WriteLine(string text, ConsoleColor color)
        {
            ConsoleColor prevColor = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.WriteLine(text);
            Console.ForegroundColor = prevColor;
        }

        public static void WriteLine()
        {
            Console.WriteLine();
        }

        public static string ReadLine()
        {
            return Console.ReadLine();
        }
    }
}
