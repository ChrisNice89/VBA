using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using Skynet.DAO.Reflection;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {

            New<Test>.Compile();
           
            Stopwatch benchmark = Stopwatch.StartNew();
            for (int i = 0; i < 100000; i++)
            {
                var result = New<Test>.Instance();//Type<Test>.New();
            }
            benchmark.Stop();
            Console.WriteLine(benchmark.ElapsedMilliseconds + " Type<Test>.New()");

            benchmark = Stopwatch.StartNew();
            for (int i = 0; i < 100000; i++)
            {
                System.Activator.CreateInstance<Test>();
            }
            benchmark.Stop();
            Console.WriteLine(benchmark.ElapsedMilliseconds + " System.Activator.CreateInstance<Test>();");
            Console.Read();
        }
    }

    public class Test
    {
        public Test()
        {

        }

        public Test(string name)
        {
            Name = name;
        }

        public string Name { get; set; }
    }
}
