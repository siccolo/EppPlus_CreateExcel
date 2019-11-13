using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NET_Excel.Helpers
{
    public static class NumberExtensions
    {
        //ToColumnLetter
        public static string ToColumnLetter (this int number)
        {
            if (number < 1) return String.Empty;
            return ConvertNumberToExcelColumnLetter(number);
            /*
            char c = (Char)(64 + number);
            return c.ToString();
             * */
        }

        public static string ConvertNumberToExcelColumnLetter(int column)
        {
            if (column < 1) return String.Empty;
            return ConvertNumberToExcelColumnLetter((column - 1) / 26) + (char)('A' + (column - 1) % 26);
        }
    }
}
