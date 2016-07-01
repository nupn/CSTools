using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ArrayCompareTest
{
    class Program
    {
        static string ConvertStringArrayToString(string[] arr)
        {
            StringBuilder builder = new StringBuilder();
            foreach (string value in arr)
            {
                builder.Append(value);
            }

            return builder.ToString();
        }

        static void Main(string[] args)
        {
            String[] a1 = new String[] { "aaa", "bbbc" };
            String[] a2 = new String[]{ "aaa", "bbb" };

            Console.WriteLine(ConvertStringArrayToString(a1).GetHashCode());
            Console.WriteLine(ConvertStringArrayToString(a2).GetHashCode());
        }
    }
}
