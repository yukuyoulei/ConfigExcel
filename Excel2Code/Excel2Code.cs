using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace Excel2Code
{
    public static class Excel2Code
    {
        public static void GenerateFromDir(string dir, string[] ignores)
        {
            var fs = Directory.GetFiles(dir);
            foreach (var f in fs)
            {
                var finfo = new FileInfo(f);
                string fileExt = finfo.Extension.ToLower();
                if (finfo.Name.StartsWith("~"))
                    continue;
                if (ignores != null && ignores.Contains(finfo.Name))
                {
                    Console.WriteLine($"Skip {finfo.Name}");
                    continue;
                }
                if (fileExt == "xls" || fileExt == ".xlsx")
                    Excel2Code.Generate(f);
            }
        }

        private static List<string> subClasses = new List<string>();
        public static void Generate(string file)
        {
            Console.WriteLine($"Generating {file}");
            IWorkbook workbook;
            var finfo = new FileInfo(file);
            string fileExt = finfo.Extension.ToLower();
            var classname = finfo.Name.Substring(0, finfo.Name.Length - fileExt.Length);
            if (!classname.StartsWith("Config_"))
                classname = "Config_" + classname;
            var classes = "";
            var res = "";
            res = Utils.AppendLine(res, "//本文件为自动生成，请勿手动修改");
            res = Utils.AppendLine(res, "//--------------------------");
            res = Utils.AppendLine(res, "//https://github.com/yukuyoulei/Excel2CSharp");
            res = Utils.AppendLine(res, "//--------------------------");
            res = Utils.AppendLine(res, "//");
            res = Utils.AppendLine(res, "using System.Collections.Generic;");
            res = Utils.AppendLine(res, $"public partial class {classname}");
            res = Utils.AppendLine(res, "{");
            using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                if (fileExt == "xls")
                    workbook = new HSSFWorkbook(fs);
                else if (fileExt == ".xlsx")
                    workbook = new XSSFWorkbook(fs);
                else
                    throw new Exception($"Cannot parse {file}, only support xls/xlsx files.");
                for (var i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var sheet = workbook.GetSheetAt(i);
                    var name = sheet.SheetName;
                    if (name.StartsWith("#")) // 参数表
                    {
                        res = Utils.AppendLine(res, GenerateDesc(sheet));
                    }
                    else
                    {
                        var subclass = GenerateSubDataClass(sheet);
                        if (!string.IsNullOrEmpty(subclass))
                            classes = Utils.AppendLine(classes, subclass);
                        var headcount = sheet.LastRowNum - sheet.FirstRowNum;
                        if (headcount == 3)
                            res = Utils.AppendLine(res, GenerateSubData(sheet));
                        else if (headcount > 3)
                            res = Utils.AppendLine(res, GenerateSubDataList(sheet));
                    }
                }
            }
            res = Utils.AppendLine(res, "}");
            res = Utils.AppendLine(res, classes);

            var path = $"Codes/{classname}.cs";
            File.WriteAllText(path, res);
            Console.WriteLine($"Generated {path}");
        }


        private static string GenerateSubDataClass(ISheet sheet)
        {
            var classname = sheet.SheetName.Trim().Split(' ')[0];
            if (subClasses.Contains(classname))
                return "";
            subClasses.Add(classname);
            var res = "";
            res = Utils.AppendLine(res, $"public class {classname}");
            res = Utils.AppendLine(res, "{");
            var rowComment = sheet.GetRow(0);
            var rowType = sheet.GetRow(1);
            var rowParamName = sheet.GetRow(2);
            for (var i = 0; i < rowType.LastCellNum; i++)
            {
                var paramName = rowParamName?.GetCell(i)?.ToString();
                if (paramName == "ignore")
                    continue;
                if (string.IsNullOrEmpty(paramName))
                    break;
                var realtype = GetRealType(rowType.GetCell(i));
                res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}public {realtype} {paramName}; //{rowComment.GetCell(i)}");
            }
            res = Utils.AppendLine(res, "}");
            return res;
        }

        private static string GenerateSubDataList(ISheet sheet)
        {
            var classname = sheet.SheetName.Trim();
            var asheetname = classname.Split(' ');
            var paramname = $"l{classname}";
            if (asheetname.Length >= 2)
            {
                classname = asheetname[0];
                paramname = asheetname[1];
            }
            var res = "";

            var paramNameRow = sheet.GetRow(2);
            var columns = paramNameRow.LastCellNum;
            var rowType = sheet.GetRow(1);
            var rowComment = sheet.GetRow(0);
            var firstRowType = rowType.GetCell(0).ToString();
            var genericList = rowComment.GetCell(0).ToString().ToLower().Contains("list");
            var allres = "";
            if (genericList)
            {
                res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}public static List<{classname}> {paramname} = new List<{classname}>()");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}{{");
                for (int i = 3; i <= sheet.LastRowNum; i++)
                {
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}new {classname}");
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}{{");
                    var row = sheet.GetRow(i);
                    for (var j = 0; j < columns; j++)
                    {
                        var realValue = CellToString(row.GetCell(j), rowType.GetCell(j).ToString());
                        if (string.IsNullOrEmpty(realValue))
                            continue;
                        var paramName = paramNameRow.GetCell(j).ToString();
                        if (string.IsNullOrEmpty(paramName))
                            break;
                        res = Utils.AppendLine(res, $"{Utils.GetTabs(3)}{paramName} = {realValue},");
                    }
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}}},");
                }

            }
            else
            {
                if (asheetname.Length < 2)
                    paramname = $"d{classname}";
                var keyType = rowType.GetCell(0).ToString();
                var keyParamName = paramNameRow.GetCell(0).ToString();

                res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}public static Dictionary<{keyType}, {classname}> {paramname} = new Dictionary<{keyType}, {classname}>();");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}public static {classname} OnGetFrom_{paramname}({keyType} {keyParamName})");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}{{");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}switch ({keyParamName})");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}{{");
                for (int i = 3; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    var r0 = row.GetCell(0);
                    if (r0 == null)
                        break;
                    var srow0 = r0.ToString();
                    if (string.IsNullOrEmpty(srow0))
                        break;
                    if (firstRowType == "string")
                        srow0 = $"\"{srow0}\"";
                    allres += srow0 + ",";
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(3)}case {srow0}:");
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(4)}if (!{paramname}.ContainsKey({srow0}))");
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(4)}{{");
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(5)}var data = new {classname}()");
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(5)}{{");
                    for (var j = 0; j < columns; j++)
                    {
                        var realtype = rowType.GetCell(j).ToString();
                        var realValue = CellToString(row.GetCell(j), realtype);
                        if (string.IsNullOrEmpty(realValue))
                            continue;
                        var paramName = paramNameRow.GetCell(j).ToString();
                        if (string.IsNullOrEmpty(paramName))
                            break;
                        if (paramName == "ignore")
                            continue;
                        res = Utils.AppendLine(res, $"{Utils.GetTabs(6)}{paramName} = {realValue},");
                    }
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(5)}}};");
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(5)}{paramname}[{srow0}] = data;");
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(4)}}}");
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(4)}return {paramname}[{srow0}];");
                }
                res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}}}");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}return null;");
            }
            res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}}}");
            if (!string.IsNullOrEmpty(allres))
                res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}public static List<{firstRowType}> all{classname}s = new List<{firstRowType}>(){{{allres}}};");
            return res;
        }

        private static string GenerateSubData(ISheet sheet)
        {
            var classname = sheet.SheetName.Trim();
            var asheetname = classname.Split(' ');
            var paramname = $"m{classname}";
            if (asheetname.Length == 2)
            {
                classname = asheetname[0];
                paramname = asheetname[1];
            }
            var res = "";
            res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}public static {classname} {paramname} = new {classname}()");
            res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}{{");
            var paramNameRow = sheet.GetRow(2);
            var columns = paramNameRow.LastCellNum;
            var rowType = sheet.GetRow(1);
            for (int i = 3; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                for (var j = 0; j < columns; j++)
                {
                    var realValue = CellToString(row.GetCell(j), rowType.GetCell(j).ToString());
                    if (string.IsNullOrEmpty(realValue))
                        continue;
                    var paramName = paramNameRow.GetCell(j).ToString();
                    if (string.IsNullOrEmpty(paramName))
                        break;
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}{paramName} = {realValue},");
                }
            }
            res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}}};");
            return res;
        }

        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell">目标单元格</param>
        /// <param name="OnlyType">是否只判断cell的类型</param>
        /// <returns></returns>
        private static string GetValueType(ICell cell, bool OnlyType)
        {
            if (cell == null)
                throw new Exception($"Missing cell");
            switch (cell.CellType)
            {
                case CellType.Boolean:
                    return "bool";
                case CellType.Numeric:
                    return "int";
                case CellType.Blank:
                    return "string";
                case CellType.String:
                case CellType.Error:
                case CellType.Formula:
                default:
                    return OnlyType ? "string" : cell.ToString();
            }
        }
        private static string GenerateDesc(ISheet sheet)
        {
            var res = "";
            for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                var valueType = row.GetCell(0);
                var valueName = row.GetCell(1).ToString();
                var value = row.GetCell(2);
                var realType = valueType == null ? GetValueType(value, true) : valueType.ToString();
                var realValue = CellToString(value, realType);
                if (realValue == null)
                    return res;
                if (string.IsNullOrEmpty(valueName))
                    return res;
                res = Utils.AppendLine(res,
                    $"{Utils.GetTabs(1)}public static {realType} {valueName} = {realValue};"
                    );
            }
            return res;
        }
        private static Dictionary<string, string> drepalace = new Dictionary<string, string>()
        {
            {"map", "Dictionary" },
            {"v2", "DataVector2" },
            {"v3", "DataVector3" },
        };
        private static string GetRealType(ICell cell, string realType = "")
        {
            realType = string.IsNullOrEmpty(realType) ? GetValueType(cell, false) : realType;
            foreach (var kv in drepalace)
                realType = realType.Replace(kv.Key, kv.Value);
            return realType;
        }
        private static string CellToString(ICell cell, string realType)
        {
            if (cell == null)
                return null;
            var realValue = cell.ToString();
            realType = GetRealType(cell, realType);
            if (realType == "bool")
                realValue = realValue.ToLower();
            else if (realType == "string" && !realValue.EndsWith("\""))
                realValue = $"\"{realValue.Replace("\\", "\\\\")}\"";
            else if (realType.EndsWith("[]"))
            {
                if (!string.IsNullOrEmpty(realValue) && realType.StartsWith("float"))
                {
                    string pattern = @"\d+\.\d+"; // regular expression pattern to match numbers
                    string replacement = "$0f"; // replacement string with $0 representing the entire match and "f" after it

                    realValue = Regex.Replace(realValue, pattern, replacement);
                }
                if (!realValue.StartsWith("["))
                    realValue = $"new {realType}{{{realValue}}}";
                else if (realType.EndsWith("[][]"))
                {
                    var a = realValue.Replace("[", "{")
                        .Replace("]", "}")
                        .Replace("{", "new []{");
                    realValue = $"new {realType}{{{a}}}";
                }
                else
                    realValue = $"new {realType}{realValue.Replace("[", "{").Replace("]", "}")}";
            }
            else if (realType == "float")
                realValue = $"{realValue}{(!realValue.EndsWith("f") ? "f" : "")}";
            else if (realType.StartsWith("Dictionary<"))
                if (realValue.StartsWith("{{"))
                    realValue = $"new {realType}(){realValue}";
                else
                    realValue = $"new {realType}(){{{realValue}}}";
            else if (realType.StartsWith("Data"))
            {
                var res = $"new {realType}().FromString(\"{realValue}\")";
                realValue = res;
            }
            return realValue;
        }
    }

}
