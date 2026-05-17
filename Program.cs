using System;
using System.Windows.Forms;
using System.IO;
using System.Text;

namespace CaesarProject
{
    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                AttachConsole(-1);
                RunConsoleMode(args);
                return;
            }

            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }

        static void RunConsoleMode(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Использование: Caesar.exe [e/d] [сдвиг] [текст или путь к файлу]");
                Console.WriteLine("Пример: Caesar.exe e 3 \"Hello\"");
                Console.WriteLine("Пример с файлом: Caesar.exe e 3 \"C:\\docs\\file.txt\"");
                return;
            }

            string mode = args[0];
            int shift = int.Parse(args[1]);
            string input = args[2];

            // Проверяем — файл или текст
            string text;
            bool isFile = File.Exists(input);

            if (isFile)
            {
                text = File.ReadAllText(input, Encoding.UTF8);
            }
            else
            {
                text = input;
            }

            int actualShift = mode == "d" ? -shift : shift;
            string result = ApplyCaesar(text, actualShift);

            if (isFile)
            {
                File.WriteAllText(input, result, Encoding.UTF8);
                Console.WriteLine("Готово! Файл сохранён: " + input);
            }
            else
            {
                Console.WriteLine("Результат: " + result);
            }
        }

        static string ApplyCaesar(string input, int shift)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            StringBuilder res = new StringBuilder();
            foreach (char c in input)
            {
                if (c >= 'a' && c <= 'z') res.Append((char)('a' + (c - 'a' + (shift % 26) + 26) % 26));
                else if (c >= 'A' && c <= 'Z') res.Append((char)('A' + (c - 'A' + (shift % 26) + 26) % 26));
                else if (c >= 'а' && c <= 'я') res.Append((char)('а' + (c - 'а' + (shift % 32) + 32) % 32));
                else if (c >= 'А' && c <= 'Я') res.Append((char)('А' + (c - 'А' + (shift % 32) + 32) % 32));
                else res.Append(c);
            }
            return res.ToString();
        }

        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        static extern bool AttachConsole(int dwProcessId);
    }
}
