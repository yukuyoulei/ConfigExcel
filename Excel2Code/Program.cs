using System;
using System.IO;

namespace Excel2Code
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var outdir = "./exportcsharp";
                var compiledir = "./exportcsharp";
                if (args.Length == 0)
                {
                    Console.WriteLine("Input a dir:");
                    var dir = Console.ReadLine();
                    Excel2Code.GenerateFromDir(dir, outdir, compiledir, new string[0]);
                }
                else if (args[0] == "-dir")
                {
                    outdir = "";
                    compiledir = "";
                    var ignores = "";
                    for (var i = 1; i < args.Length; i++)
                    {
                        if (args[i] == "-out")
                            outdir = args[i + 1];
                        else if (args[i] == "-compile")
                            compiledir = args[i + 1];
                        else if (args[i] == "-ignore")
                            ignores = args[i + 1];
                    }
                    Excel2Code.GenerateFromDir(args[1], outdir, compiledir, ignores.Split(','));
                }
                else if (args[0] == "-autoc")
                {

                }
                else
                    foreach (var f in args)
                    {
                        Excel2Code.Generate(f, outdir);
                    }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"导出失败！\r\n"
                    + $"出错的文件：{Excel2Code.ExportingFile}\r\n"
                    + $"出错的Sheet：{Excel2Code.ExportingSheet}\r\n"
                    + $"{ex}");
                Console.ReadKey();
            }
        }
    }
}
