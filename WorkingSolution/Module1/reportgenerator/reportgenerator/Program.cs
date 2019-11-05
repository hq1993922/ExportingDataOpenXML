using System;

namespace reportgenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Generator g = new Generator();
            g.CreatePackage(args[0]);
            Console.ReadLine();
        }
    }
}
