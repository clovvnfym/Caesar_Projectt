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

            string mode = args[0].ToLower();
            if (mode != "e" && mode != "d")
            {
                Console.WriteLine("Ошибка: первый аргумент должен быть 'e' (зашифровать) или 'd' (расшифровать).");
                return;
            }

            if (!int.TryParse(args[1], out int shift))
            {
                Console.WriteLine("Ошибка: сдвиг должен быть целым числом (например, 3 или -5).");
                return;
            }

            string input = args[2];
            bool isFile = File.Exists(input);
            string text;

            if (isFile)
            {
                if (Path.GetExtension(input).ToLower() == ".docx")
                {
                    Console.WriteLine("Ошибка: файлы .docx не поддерживаются в консольном режиме.");
                    Console.WriteLine("Используйте графический интерфейс для работы с .docx файлами.");
                    return;
                }
                text = File.ReadAllText(input, Encoding.UTF8);
            }
            else
            {
                text = input;
            }

            int actualShift = mode == "d" ? -shift : shift;

            string result = CaesarCipher.Apply(text, actualShift);

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
        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        static extern bool AttachConsole(int dwProcessId);
    }
}
