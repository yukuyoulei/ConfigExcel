using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using UnityEngine;

namespace Excel2Code
{
    public static class Excel2Code
    {
        private static string OutputDir;
        public static void GenerateFromDir(string dir, List<string> lignore, string outDir)
        {
            OutputDir = outDir;
            if (!Directory.Exists(outDir))
            {
                Directory.CreateDirectory(outDir);
            }
            var fs = Directory.GetFiles(dir);
            foreach (var f in fs)
            {
                var finfo = new FileInfo(f);
                string fileExt = finfo.Extension.ToLower();
                if (finfo.Name.StartsWith("~"))
                {
                    Debug.Log($"跳过{finfo.FullName}");
                    continue;
                }
                if (lignore.Contains(finfo.Name))
                {
                    Debug.Log($"跳过{finfo.FullName}");
                    continue;
                }
                try
                {
                    if (fileExt == "xls" || fileExt == ".xlsx")
                        Excel2Code.Generate(f);
                }
                catch (Exception ex)
                {
                    Debug.Log($"导表出错 {f}\r\n{ex}");
                }
            }
        }

        private static List<string> subClasses = new List<string>();
        public static void Generate(string file)
        {
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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
                using (var package = new ExcelPackage(fs))
                {
                    var workbook = package.Workbook;
                    for (var i = 0; i < workbook.Worksheets.Count; i++)
                    {
                        var sheet = workbook.Worksheets[i + 1]; // EPPlus is 1-based
                        var name = sheet.Name;
                        if (name.StartsWith("#")) // 参数表
                        {
                            res = Utils.AppendLine(res, GenerateDesc(sheet));
                        }
                        else
                        {
                            var subclass = GenerateSubDataClass(sheet);
                            if (!string.IsNullOrEmpty(subclass))
                                classes = Utils.AppendLine(classes, subclass);
                            var headcount = sheet.Dimension.End.Row - 1;
                            if (headcount == 3)
                                res = Utils.AppendLine(res, GenerateSubData(sheet));
                            else if (headcount > 3)
                                res = Utils.AppendLine(res, GenerateSubDataList(sheet));
                        }
                    }
                }
            }
            res = Utils.AppendLine(res, "}");
            res = Utils.AppendLine(res, classes);

            var path = $"{OutputDir}/{classname}.cs";
            File.WriteAllText(path, res);
            Debug.Log($"Generated {path}");
        }

        private static string GenerateSubDataClass(ExcelWorksheet sheet)
        {
            var classname = sheet.Name.Trim().Split(' ')[0];
            if (subClasses.Contains(classname))
                return "";
            subClasses.Add(classname);
            var res = "";
            res = Utils.AppendLine(res, $"public partial class {classname}");
            res = Utils.AppendLine(res, "{");

            var rowComment = sheet.Cells[1, 1, 1, sheet.Dimension.End.Column];
            var rowType = sheet.Cells[2, 1, 2, sheet.Dimension.End.Column];
            var rowParamName = sheet.Cells[3, 1, 3, sheet.Dimension.End.Column];

            for (var i = 1; i <= sheet.Dimension.End.Column; i++)
            {
                var paramName = rowParamName[3, i].Text;
                if (paramName == "ignore")
                    continue;
                if (string.IsNullOrEmpty(paramName))
                    break;
                res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}public {GetRealType(rowType[2, i])} {paramName}; //{rowComment[1, i].Text}");
            }
            res = Utils.AppendLine(res, "}");
            return res;
        }

        private static string GenerateSubDataList(ExcelWorksheet sheet)
        {
            var classname = sheet.Name.Trim();
            var asheetname = classname.Split(' ');
            var paramname = $"l{classname}";
            if (asheetname.Length >= 2)
            {
                classname = asheetname[0];
                paramname = asheetname[1];
            }
            var res = "";

            var paramNameRow = sheet.Cells[3, 1, 3, sheet.Dimension.End.Column];
            var columns = sheet.Dimension.End.Column;
            var rowType = sheet.Cells[2, 1, 2, columns];
            var rowComment = sheet.Cells[1, 1, 1, columns];
            var firstRowType = rowType[2, 1].Text;
            var genericList = rowComment[1, 1].Text.ToLower().Contains("list");
            var allres = "";

            if (genericList)
            {
                res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}public static List<{classname}> {paramname} = new List<{classname}>()");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}{{");
                for (int i = 4; i <= sheet.Dimension.End.Row; i++)
                {
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}new {classname}");
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}{{");
                    var row = sheet.Cells[i, 1, i, columns];
                    for (var j = 1; j <= columns; j++)
                    {
                        var realValue = CellToString(row[i, j], rowType[2, j].Text);
                        if (string.IsNullOrEmpty(realValue))
                            continue;
                        var paramName = paramNameRow[3, j].Text;
                        if (paramName == "ignore")
                            continue;
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
                var keyType = rowType[2, 1].Text;
                var keyParamName = paramNameRow[3, 1].Text;

                res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}public static Dictionary<{keyType}, {classname}> {paramname} = new Dictionary<{keyType}, {classname}>();");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}public static {classname} OnGetFrom_{paramname}({keyType} {keyParamName})");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}{{");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}switch ({keyParamName})");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}{{");

                for (int i = 4; i <= sheet.Dimension.End.Row; i++)
                {
                    var cell = sheet.Cells[i, 1];
                    if (cell == null || string.IsNullOrEmpty(cell.Text))
                        break;

                    var srow0 = cell.Text;
                    if (firstRowType == "string")
                        srow0 = $"\"{srow0}\"";
                    allres += srow0 + ",";

                    res = Utils.AppendLine(res, $"{Utils.GetTabs(3)}case {srow0}:");
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(4)}if (!{paramname}.ContainsKey({srow0}))");
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(4)}{{");
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(5)}var data = new {classname}()");
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(5)}{{");

                    var row = sheet.Cells[i, 1, i, columns];
                    for (var j = 1; j <= columns; j++)
                    {
                        var realValue = CellToString(row[i, j], rowType[2, j].Text);
                        if (string.IsNullOrEmpty(realValue))
                            continue;
                        var paramName = paramNameRow[3, j].Text;
                        if (paramName == "ignore")
                            continue;
                        if (string.IsNullOrEmpty(paramName))
                            break;
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

        private static string GenerateSubData(ExcelWorksheet sheet)
        {
            var classname = sheet.Name.Trim();
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

            var paramNameRow = sheet.Cells[3, 1, 3, sheet.Dimension.End.Column];
            var columns = sheet.Dimension.End.Column;
            var rowType = sheet.Cells[2, 1, 2, columns];

            for (int i = 4; i <= sheet.Dimension.End.Row; i++)
            {
                var row = sheet.Cells[i, 1, i, columns];
                for (var j = 1; j <= columns; j++)
                {
                    var realValue = CellToString(row[i, j], rowType[2, j].Text);
                    if (string.IsNullOrEmpty(realValue))
                        continue;
                    var paramName = paramNameRow[3, j].Text;
                    if (paramName == "ignore")
                        continue;
                    if (string.IsNullOrEmpty(paramName))
                        break;
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}{paramName} = {realValue},");
                }
            }
            res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}}};");
            return res;
        }

        private static string GetValueType(ExcelRange cell, bool OnlyType)
        {
            if (cell == null)
                throw new Exception($"Missing cell");

            if (cell.Value == null)
                return "string";

            switch (cell.Value.GetType().Name.ToLower())
            {
                case "boolean":
                    return "bool";
                case "double":
                case "int32":
                    return "int";
                default:
                    return OnlyType ? "string" : cell.Text;
            }
        }

        private static string GenerateDesc(ExcelWorksheet sheet)
        {
            var res = "";
            for (int i = 1; i <= sheet.Dimension.End.Row; i++)
            {
                var valueType = sheet.Cells[i, 1];
                var valueName = sheet.Cells[i, 2].Text;
                var value = sheet.Cells[i, 3];
                var realType = valueType.Value == null ? GetValueType(value, true) : valueType.Text;
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

        private static string GetRealType(ExcelRange cell, string realType = "")
        {
            realType = string.IsNullOrEmpty(realType) ? GetValueType(cell, false) : realType;
            foreach (var kv in drepalace)
                realType = realType.Replace(kv.Key, kv.Value);
            return realType;
        }

        private static string CellToString(ExcelRange cell, string realType)
        {
            if (cell == null || cell.Value == null)
                return null;
            var realValue = cell.Text;
            realType = GetRealType(cell, realType);
            if (realType == "bool")
                realValue = realValue == "1" ? "true" : "false";
            else if (realType == "string" && !realValue.EndsWith("\""))
                realValue = $"@\"{realValue}\"";
            else if (realType.EndsWith("[]"))
            {
                if (!string.IsNullOrEmpty(realValue) && realType.StartsWith("float"))
                {
                    string pattern = "\\d+\\.\\d+";
                    string replacement = "$0f";
                    realValue = Regex.Replace(realValue, pattern, replacement);
                }
                if (realValue.Contains("["))
                {
                    if (realType.EndsWith("[][]"))
                    {
                        realValue = $"new {realType}{{{realValue.Replace("]", "}").Replace("[", "new[]{")}}}";
                    }
                    else
                        realValue = $"new {realType}{realValue.Replace("[", "{").Replace("]", "}")}";
                }
                else
                    realValue = $"new {realType}{{{realValue}}}";
            }
            else if (realType == "float")
            {
                if (string.IsNullOrEmpty(realValue))
                    return null;
                realValue = $"{realValue}{(!realValue.EndsWith("f") ? "f" : "")}";
            }
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