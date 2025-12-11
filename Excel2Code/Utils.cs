using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Code
{
    public static class Utils
    {
        private static readonly string[] tabCache = new string[10];
        
        static Utils()
        {
            // Pre-cache tab strings for better performance
            for (int i = 0; i < tabCache.Length; i++)
            {
                tabCache[i] = new string('\t', Math.Max(0, i - 1));
            }
        }
        
        public static string GetTabs(int count)
        {
            if (count < tabCache.Length)
                return tabCache[count];
                
            // For larger counts, create on-demand
            return new string('\t', Math.Max(0, count - 1));
        }
        
        public static string AppendLine(string res, string value)
        {
            return $"{res}{value}\n";
        }
    }
}