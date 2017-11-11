using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelEditor
{
    public  static class Helper
    {
        public static string NextLetter(string letter)
        {
            //Caps check
            if (!IsAllUpper(letter))
                return "";

            const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            if (!String.IsNullOrEmpty(letter))
            {
                char lastLetterInString = letter[letter.Length - 1];

                // if the last letter in the string is the last letter of the alphabet
                if (alphabet.IndexOf(lastLetterInString) == alphabet.Length - 1)
                {
                    //replace the last letter in the string with the first leter of the alphabet and get the next letter for the rest of the string
                    return NextLetter(letter.Substring(0, letter.Length - 1)) + alphabet[0];
                }
                else
                {
                    // replace the last letter in the string with the proceeding letter of the alphabet
                    return letter.Remove(letter.Length - 1).Insert(letter.Length - 1, (alphabet[alphabet.IndexOf(letter[letter.Length - 1]) + 1]).ToString());
                }
            }
            //return the first letter of the alphabet
            return alphabet[0].ToString();
        }

        public static bool IsAllUpper(string input)
        {
            for (int i = 0; i < input.Length; i++)
            {
                if (!Char.IsUpper(input[i]))
                    return false;
            }

            return true;
        }
    }
}
