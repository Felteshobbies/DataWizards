using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataWizard;

namespace test
{
    internal class Program
    {
        static void Main(string[] args)
        {
            CSV wiz = new CSV();
            wiz.Load(@"C:\temp\test2.txt");

            Console.WriteLine(wiz.Separator);
            Console.WriteLine(wiz.SeparatorProbability);
            Console.WriteLine(wiz.StartLine);
            Console.ReadKey();
        }
    }
}
