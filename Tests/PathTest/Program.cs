using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PathTest
{
    class Program
    {
        static void Main(string[] args)
        {

            String path = System.IO.Path.GetFullPath("../../test.xlsx");
            Console.WriteLine(path);

            Console.WriteLine("PathTest");
            Console.ReadKey();
        }
    }
}
