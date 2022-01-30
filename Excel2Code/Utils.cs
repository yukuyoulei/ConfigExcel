using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Code
{
	public static class Utils
	{
		public static string GetTabs(int count)
		{
			var res = "";
			for (var i = 0;i<count;i++)
			{
				res += "\t";
			}
			return res;
		}
		public static string AppendLine(string res, string value)
		{
			return $"{res}{value}\r\n";
		}
	}
}
