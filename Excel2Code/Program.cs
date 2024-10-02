using System;
using System.IO;

namespace Excel2Code
{
    class Program
    {
        static void Main(string[] args)
        {
            if (!Directory.Exists("Codes"))
                Directory.CreateDirectory("Codes");
            string[] ignores = null;
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "ignore")
                {
                    ignores = args[++i].Split(',');
                    break;
                }
            }
            if (args.Length == 0)
            {
                Console.WriteLine("Input a dir:");
                var dir = Console.ReadLine();
                Excel2Code.GenerateFromDir(dir, ignores);
            }
            else if (args[0] == "-dir")
            {
                Excel2Code.GenerateFromDir(args[1], ignores);
            }
            else
                foreach (var f in args)
                {
                    Excel2Code.Generate(f);
                }
            //Console.WriteLine("Press any key to exit...");
            //Console.ReadKey();
        }
    }
}
