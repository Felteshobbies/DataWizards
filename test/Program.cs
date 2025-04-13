using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataWizard;
using libDataWizard;

namespace test
{
    internal class Program
    {
        static void Main(string[] args)
        {
            CSV wiz = new CSV();
            wiz.Load(@"C:\temp\test3.txt");

            Console.WriteLine(wiz.Separator);
            Console.WriteLine(wiz.SeparatorProbability);
            Console.WriteLine(wiz.StartLine);
            XLS xls = new XLS(@"c:\temp\test.xlsx", true);



            Console.ReadKey();



        }
    }
}
