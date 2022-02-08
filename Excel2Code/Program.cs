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
			if (args.Length == 0)
			{
				Console.WriteLine("Input a dir:");
				var dir = Console.ReadLine();
				Excel2Code.GenerateFromDir(dir);
			}
			else if (args[0] == "-dir")
			{
				Excel2Code.GenerateFromDir(args[1]);
			}
			else
				foreach (var f in args)
				{
					Excel2Code.Generate(f);
				}
			Console.WriteLine("Press any key to exit...");
			Console.ReadKey();
		}
	}
}
