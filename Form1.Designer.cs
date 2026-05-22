using System;
using System.Text;

namespace CaesarProject
{
    /// <summary>
    /// Общий класс шифра Цезаря.
    /// Вынесен сюда, чтобы не дублировать логику в Form1.cs и Program.cs.
    /// </summary>
    public static class CaesarCipher
    {
        /// <summary>
        /// Применяет шифр Цезаря к строке.
        /// Поддерживает латиницу, кириллицу (без ё/Ё), цифры и знаки препинания.
        /// Неизвестные символы пропускаются без изменений.
        /// </summary>
        public static string Apply(string input, int shift)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            StringBuilder res = new StringBuilder();
            foreach (char c in input)
            {
                if      (c >= 'a' && c <= 'z') res.Append((char)('a' + (c - 'a' + (shift % 26) + 26) % 26));
                else if (c >= 'A' && c <= 'Z') res.Append((char)('A' + (c - 'A' + (shift % 26) + 26) % 26));
                else if (c >= 'а' && c <= 'я') res.Append((char)('а' + (c - 'а' + (shift % 32) + 32) % 32));
                else if (c >= 'А' && c <= 'Я') res.Append((char)('А' + (c - 'А' + (shift % 32) + 32) % 32));
                else res.Append(c); // цифры, пробелы, пунктуация, ё/Ё и прочее — без изменений
            }
            return res.ToString();
        }
    }
}
