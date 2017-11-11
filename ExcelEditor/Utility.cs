using System;
using System.Collections.Generic;

namespace ExcelEditor
{
    public static class Utility
    {
        // From: https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformat(v=office.15).aspx
        public static Dictionary<int, string> excelNumFmtsDic = new Dictionary<int, string>() 
        {
            {14, "MM-dd-yy"},
            {15, "d-MMM-yy"},
            {16, "d-MMM"},
            {17, "MMM-yy"},
        };

        public static List<char> validNumFmtsChars = new List<char> { 'd', 'D', 'm', 'M', 'y', 'Y', ',' };

        public static bool ParseAndAddNumFmt(string numFmts)
        {
            string retStirng = string.Empty;
            char? prevchar = null;
            char[] chars = numFmts.ToCharArray();

            foreach(var ch in chars)
            {
                if (prevchar != ch && prevchar != null && validNumFmtsChars.Contains(ch))
                {
                    retStirng += ' ';
                }
                if (validNumFmtsChars.Contains(ch))
                {
                    if(ch == 'm')
                    {
                        retStirng += "M";
                    }
                    else
                    {
                        retStirng += ch;
                    }
                    prevchar = ch;
                }
                
            }

            //Test
            try
            {
                double dateVal = 42395.0;
                var dateTmp = DateTime.FromOADate(dateVal);
                string dateString = dateTmp.ToString(retStirng);
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
            }
            

            return true;
        }
    }

    
}
