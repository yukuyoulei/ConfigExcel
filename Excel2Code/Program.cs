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
			if (args[0] == "-dir")
			{
				var fs = Directory.GetFiles(args[1]);
				foreach (var f in fs)
				{
					var finfo = new FileInfo(f);
					string fileExt = finfo.Extension.ToLower();
					if (fileExt == "xls" || fileExt == ".xlsx")
						Excel2Code.Generate(f);
				}
			}
			else
				foreach (var f in args)
				{
					Excel2Code.Generate(f);
				}
		}
	}
}
