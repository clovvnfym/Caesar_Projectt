using System;
using System.Text;

namespace CaesarProject
{
   
    public static class CaesarCipher
    {
   
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
                else res.Append(c); 
            }
            return res.ToString();
        }
    }
}
