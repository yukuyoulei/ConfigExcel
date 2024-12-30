using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using NPOI.SS.Formula.Functions;

namespace Excel2Code
{
    public static class Excel2Code
    {
        public static string ExportingFile;
        public static string ExportingSheet;
        public static string ExportingColumn;
        private const string historyFile = "history.txt";
        private static Dictionary<string, long> _excelHistories;
        public static Dictionary<string, long> excelHistories
        {
            get
            {
                if (_excelHistories == null)
                {
                    _excelHistories = new();
                    if (File.Exists(historyFile))
                    {
                        var all = File.ReadAllLines(historyFile);
                        foreach (var line in all)
                        {
                            var aline = line.Trim().Split(',');
                            if (aline.Length != 2)
                                continue;
                            _excelHistories[aline[0]] = long.Parse(aline[1]);
                        }
                    }
                }
                return _excelHistories;
            }
        }
        private static void SaveHistory()
        {
            var res = "";
            foreach (var kv in _excelHistories)
            {
                if (!string.IsNullOrEmpty(res))
                    res += "\r\n";
                res += kv.Key + "," + kv.Value;
            }
            File.WriteAllText(historyFile, res);
        }
        public static void GenerateFromDir(string dir, string outdir, string compiledir, string[] ignores)
        {
            if (!string.IsNullOrEmpty(outdir))
            {
                if (!Directory.Exists(outdir))
                {
                    Directory.CreateDirectory(outdir);
                }
                var fs = Directory.GetFiles(dir);
                foreach (var f in fs)
                {
                    var finfo = new FileInfo(f);
                    string fileExt = finfo.Extension.ToLower();
                    if (finfo.Name.StartsWith("~") || finfo.Name.EndsWith("~") || finfo.Name.Contains(" "))
                        continue;
                    var fname = finfo.Name.Split('.')[0];
                    if (ignores.Contains(fname))
                        continue;
                    if (fileExt == "xls" || fileExt == ".xlsx")
                    {
                        var curt = (finfo.LastWriteTime - new DateTime(1970, 1, 1)).Ticks;
                        if (excelHistories.TryGetValue(fname, out var t) && curt == t)
                        {
                            Console.WriteLine($"Already generated {f}");
                            continue;
                        }
                        Excel2Code.Generate(f, outdir);
                        excelHistories[fname] = curt;
                    }
                }
            }
            if (!string.IsNullOrEmpty(compiledir))
            {
                //获取目标目录中的所有C#文件
                List<string> cSharpFiles = Directory.GetFiles(compiledir, "*.cs").ToList();
                var dirs = Directory.GetDirectories(compiledir);
                foreach (var d in dirs)
                    cSharpFiles.AddRange(Directory.GetFiles(d, "*.cs"));
                //获取目标目录中的所有C#文件

                //设置编译器选项
                CSharpCompilationOptions options = new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary);
                //添加引用
                MetadataReference[] references = new MetadataReference[]
                {
                MetadataReference.CreateFromFile("System.Private.CoreLib.dll"),
                MetadataReference.CreateFromFile("System.Console.dll")
                };

                //编译所有C#文件
                CSharpCompilation compilation = CSharpCompilation.Create("temp", cSharpFiles.Select(x => CSharpSyntaxTree.ParseText(File.ReadAllText(x))), references, options);
                var dllpath = Path.Combine(compiledir, "temp.dll");
                EmitResult result = compilation.Emit(dllpath);

                //检查是否有编译错误
                if (result.Success)
                {
                    File.Delete(dllpath);
                    Console.BackgroundColor = ConsoleColor.DarkGreen;
                    Console.WriteLine(@"
√√√√√√√√√√√√√
√                      √
√√√   编译成功   √√√
√                      √
√√√√√√√√√√√√√
");
                    Console.BackgroundColor = ConsoleColor.Black;

                    SaveHistory();
                }
                else
                {
                    Console.BackgroundColor = ConsoleColor.Red;
                    Console.WriteLine(@"
×××××××××××××
×                      ×
×××   编译错误   ×××
×                      ×
×××××××××××××
");
                    Console.BackgroundColor = ConsoleColor.Black;
                    foreach (Diagnostic error in result.Diagnostics)
                    {
                        Console.WriteLine(error.GetMessage());
                    }
                }
            }
        }

        private static List<string> subClasses = new List<string>();
        static List<ISheet> pendingSheets = new();
        public static void Generate(string file, string outdir)
        {
            ExportingFile = file;

            IWorkbook workbook;
            var finfo = new FileInfo(file);
            string fileExt = finfo.Extension.ToLower();
            var classname = finfo.Name.Substring(0, finfo.Name.Length - fileExt.Length);
            if (!classname.Contains("Config_"))
                classname = $"Config_{classname}";
            var classes = "";
            var res = "";
            res = Utils.AppendLine(res, "//本文件为自动生成，请勿手动修改");
            res = Utils.AppendLine(res, "//--------------------------");
            res = Utils.AppendLine(res, "//https://github.com/yukuyoulei/Excel2CSharp");
            res = Utils.AppendLine(res, "//--------------------------");
            res = Utils.AppendLine(res, "//");
            res = Utils.AppendLine(res, "using System.Collections.Generic;");
            res = Utils.AppendLine(res, "namespace ConfigExcel");
            res = Utils.AppendLine(res, "{");
            res = Utils.AppendLine(res, $"\tpublic partial class {classname}");
            res = Utils.AppendLine(res, "\t{");
            using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                pendingSheets.Clear();
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
                    ExportingSheet = name;

                    if (name.StartsWith("#")) // 参数表
                    {
                        res = Utils.AppendLine(res, GenerateDesc(sheet));
                    }
                    else if (name.StartsWith("Sheet"))
                        continue;
                    else
                    {
                        if (sheet.SheetName.StartsWith("="))
                        {
                            pendingSheets.Add(sheet);
                            continue;
                        }
                        var subclass = GenerateSubDataClass(sheet);
                        if (!string.IsNullOrEmpty(subclass))
                            classes = Utils.AppendLine(classes, subclass);
                        var headcount = sheet.LastRowNum - sheet.FirstRowNum;
                        //if (headcount == 3)
                        //    res = Utils.AppendLine(res, GenerateSubData(sheet));
                        //else if (headcount > 3)
                        res = Utils.AppendLine(res, GenerateSubDataList(sheet));
                        pendingSheets.Clear();
                    }
                }
            }
            res = Utils.AppendLine(res, "\t}");
            res = Utils.AppendLine(res, classes);
            res = Utils.AppendLine(res, "}");

            var path = $"{outdir}/{classname}.cs";
            File.WriteAllText(path, res);
            Console.WriteLine($"Generated {path}");
        }


        private static string GenerateSubDataClass(ISheet sheet)
        {
            var asheetname = sheet.SheetName.Trim().Split(new char[] { ' ', '|' });
            var classname = asheetname[0];
            if (asheetname.Length == 2)
            {
                classname = asheetname[1];
            }
            else if (asheetname.Length == 3)
            {
                classname = asheetname[2] + "Data";
            }
            if (subClasses.Contains(classname))
                return "";
            subClasses.Add(classname);
            var res = "";
            res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}public partial class {classname}");
            res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}{{");
            var rowComment = sheet.GetRow(0);
            var rowParamName = sheet.GetRow(1);
            var rowType = sheet.GetRow(2);
            for (var i = 0; i < rowType.LastCellNum; i++)
            {
                var paramName = rowParamName.GetCell(i)?.ToString();
                if (string.IsNullOrEmpty(paramName))
                    break;
                var realtype = GetRealType(rowType.GetCell(i));
                if (realtype == "ignore")
                    continue;
                res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}public {realtype} {paramName}; //{rowComment.GetCell(i)}");
            }
            res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}}}");
            return res;
        }

        private static string GenerateSubDataList(ISheet sheet)
        {
            var classname = sheet.SheetName.Trim();
            var asheetname = classname.Split(new char[] { ' ', '|' });
            var paramname = $"l{classname}";
            if (asheetname.Length == 2)
            {
                classname = asheetname[1];
            }
            else if (asheetname.Length == 3)
            {
                classname = asheetname[2] + "Data";
            }
            paramname = $"m{classname}";
            var res = "";

            var paramNameRow = sheet.GetRow(1);
            var columns = paramNameRow.LastCellNum;
            var rowType = sheet.GetRow(2);
            var rowComment = sheet.GetRow(0);
            var firstRowType = GetRealType(rowType.GetCell(0));
            var genericList = rowComment.GetCell(0).ToString().ToLower().Contains("list");
            var allres = "";
            if (genericList)
            {
                res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}public static List<{classname}> {paramname} = new List<{classname}>()");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}{{");
                for (int i = 3; i <= sheet.LastRowNum; i++)
                {
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(3)}new {classname}");
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(3)}{{");
                    var row = sheet.GetRow(i);
                    for (var j = 0; j < columns; j++)
                    {
                        var realValue = CellToString(row.GetCell(j), rowType.GetCell(j).ToString());
                        if (string.IsNullOrEmpty(realValue))
                            continue;
                        var paramName = paramNameRow.GetCell(j).ToString();
                        if (string.IsNullOrEmpty(paramName))
                            break;
                        res = Utils.AppendLine(res, $"{Utils.GetTabs(4)}{paramName} = {realValue},");
                    }
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(3)}}},");
                }

            }
            else
            {
                if (asheetname.Length < 2)
                    paramname = $"d{classname}";
                var keyType = rowType.GetCell(0).ToString().Split(':')[0];
                var keyParamName = paramNameRow.GetCell(0).ToString();
                var dicres = "";
                dicres = Utils.AppendLine(dicres, $"{Utils.GetTabs(2)}public static Dictionary<{keyType}, {classname}> {paramname} = new Dictionary<{keyType}, {classname}>();");
                dicres = Utils.AppendLine(dicres, $"{Utils.GetTabs(2)}public static {classname} OnGetFrom_{paramname}({keyType} {keyParamName})");
                dicres = Utils.AppendLine(dicres, $"{Utils.GetTabs(2)}{{");
                dicres = Utils.AppendLine(dicres, $"{Utils.GetTabs(3)}switch ({keyParamName})");
                dicres = Utils.AppendLine(dicres, $"{Utils.GetTabs(3)}{{");
                var datares = "";
                datares = SheetToString(datares, firstRowType, ref allres, columns, paramname
                    , classname, rowType, paramNameRow
                    , sheet);
                foreach (var s in pendingSheets)
                {
                    datares = SheetToString(datares, firstRowType, ref allres, columns, paramname
                        , classname, rowType, paramNameRow
                        , s);
                }
                if (string.IsNullOrEmpty(datares))
                    return res;
                res += dicres;
                res += datares;
                res = Utils.AppendLine(res, $"{Utils.GetTabs(3)}}}");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(3)}return null;");
            }
            res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}}}");
            if (!string.IsNullOrEmpty(allres))
            {
                res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}public static List<{firstRowType}> all{classname}s = new List<{firstRowType}>()");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(3)}{{");
                res = Utils.AppendLine(res, $"{allres}");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(3)}}};");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}public static {classname} OnGetFrom_{classname}ByIndex(int index){{if (index < 0 || index >= all{classname}s.Count) return null; return OnGetFrom_{paramname}(all{classname}s[index]);}}");
            }
            return res;
        }

        private static string SheetToString(string res, string firstRowType, ref string allres
            , short columns, string paramname, string classname
            , IRow? rowType, IRow? paramNameRow
            , ISheet sheet)
        {
            for (int i = 3; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                    continue;
                var r0 = row.GetCell(0);
                if (r0 == null)
                    break;
                var srow0 = r0.ToString();
                if (string.IsNullOrEmpty(srow0))
                    break;
                if (firstRowType == "string")
                    srow0 = $"\"{srow0}\"";
                allres = Utils.AppendLine(allres, $"{Utils.GetTabs(4)}{srow0},");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(4)}case {srow0}:");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(5)}if (!{paramname}.ContainsKey({srow0}))");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(5)}{{");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(6)}var data = new {classname}()");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(6)}{{");
                for (var j = 0; j < columns; j++)
                {
                    var icell = rowType.GetCell(j);
                    if (icell == null)
                        continue;
                    var realtype = GetRealType(icell);
                    if (realtype == "ignore")
                        continue;

                    var realValue = CellToString(row.GetCell(j), rowType.GetCell(j).ToString());
                    if (string.IsNullOrEmpty(realValue))
                        continue;
                    var paramName = paramNameRow.GetCell(j).ToString();
                    if (string.IsNullOrEmpty(paramName))
                        break;
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(7)}{paramName} = {realValue},");
                }
                res = Utils.AppendLine(res, $"{Utils.GetTabs(6)}}};");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(6)}{paramname}[{srow0}] = data;");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(5)}}}");
                res = Utils.AppendLine(res, $"{Utils.GetTabs(5)}return {paramname}[{srow0}];");
            }
            return res;
        }

        private static string GenerateSubData(ISheet sheet)
        {
            var classname = sheet.SheetName.Trim();
            var asheetname = classname.Split(new char[] { ' ', '|' });
            var paramname = $"m{classname}";
            if (asheetname.Length == 2)
            {
                classname = asheetname[1];
            }
            else if (asheetname.Length == 3)
            {
                classname = asheetname[2] + "Data";
            }
            paramname = $"m{classname}";
            var res = "";
            res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}public static {classname} {paramname} = new {classname}()");
            res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}{{");
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
                    res = Utils.AppendLine(res, $"{Utils.GetTabs(3)}{paramName} = {realValue},");
                }
            }
            res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}}};");
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
                var realType = valueType == null ? GetValueType(value, true) : GetRealType(valueType);
                var realValue = CellToString(value, realType);
                if (realValue == null)
                    return res;
                if (string.IsNullOrEmpty(valueName))
                    return res;
                var key = "static";
                if (realType.Equals("string") || realType.Equals("int") || realType.Equals("float") || realType.Equals("long"))
                    key = "const";
                res = Utils.AppendLine(res,
                    $"{Utils.GetTabs(2)}public {key} {realType} {valueName} = {realValue};" // added
                    );
            }
            return res;
        }
        private static Dictionary<string, string> drepalace = new Dictionary<string, string>()
        {
            {"map", "Dictionary" },
            {"dic", "Dictionary" },
            {"v2[]", "DataVector2[]" },
            {"v2", "DataVector2" },
            {"v3[]", "DataVector3[]" },
            {"v3", "DataVector3" },
            {"luatable", "string" },
            {"luacode", "string" },
        };
        private static string GetRealType(ICell cell, string realType = "")
        {
            realType = string.IsNullOrEmpty(realType) ? GetValueType(cell, false) : realType;
            foreach (var kv in drepalace)
                realType = realType.Replace(kv.Key, kv.Value);
            realType = realType.Split(':')[0];
            return realType.Replace(" ", "");//过滤掉所有空格
        }
        private static string CellToString(ICell cell, string realType)
        {
            if (cell == null)
                return null;
            var realValue = GetCellText(cell);
            realType = GetRealType(cell, realType);
            if (realType == "bool")
                realValue = realValue.ToLower();
            else if (realType == "string" && !realValue.EndsWith("\""))
                realValue = $"\"{realValue.Replace("\"", "\\\"")}\"";
            else if (realType.StartsWith("Data") || realType.EndsWith("Data"))
            {
                if (string.IsNullOrEmpty(realValue))
                    return "null";
                if (realType.EndsWith("[]"))
                {
                    var classname = realType.Substring(0, realType.Length - 2);
                    var strvalue = realValue.Replace(" ", "");//过滤掉所有空格
                    strvalue = strvalue.Replace("},}", "}}");//过滤掉末尾可能},}
                    strvalue = strvalue.Replace("},{", "|");//将},{替换为|便于分割
                    strvalue = strvalue.Replace("}", "|");//将}替换为|便于分割
                    strvalue = strvalue.Replace("{", "|");//将{替换为|便于分割
                    var acells = strvalue.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
                    var res = $"new {realType}{{";
                    foreach (var acell in acells)
                    {
                        res += "new " + classname + "(){" + acell + "},";
                    }
                    res += "}";
                    realValue = res;
                }
                else
                {
                    var res = string.IsNullOrEmpty(realValue) ? "null" : $"new {realType}(){realValue}";
                    realValue = res;
                }
            }
            else if (realType.EndsWith("[]"))
            {
                if (string.IsNullOrEmpty(realValue))
                    return "null";
                if (realType.Contains("float"))
                {
                    var v = realValue.Replace("{", "").Replace("}", "");
                    var av = v.Split(',');
                    for (var i = 0; i < av.Length; i++)
                    {
                        if (!string.IsNullOrEmpty(av[i]))
                            av[i] = av[i] + "f";
                    }
                    realValue = $"new {realType}{{{string.Join(",", av)}}}";
                }
                else if (realType.EndsWith("[][][]"))
                {
                    var vv = realValue.Replace(" ", "").Replace("{{{", "{{").Replace("}}}", "}}").Replace("}},", "}}");
                    var avv = vv.Split("}}");
                    var lv = new List<string>();
                    var arrayType = realType.Replace("[]", "");
                    foreach (var str in avv)
                    {
                        var v = str.Replace(" ", "").Replace("{{", "{").Replace("}}", "}").Replace("},", "}");
                        var av = v.Split("}");
                        for (var i = 0; i < av.Length; i++)
                        {
                            if (!string.IsNullOrEmpty(av[i]))
                                av[i] = "new []" + av[i] + "}";
                        }
                        lv.Add($"new {arrayType}[][]{{{string.Join(",", av)}}}");
                    }
                    realValue = $"new {realType}{{{string.Join(",", lv)}}}";
                }
                else if (realType.EndsWith("[][]"))
                {
                    var v = realValue.Replace(" ", "").Replace("{{", "{").Replace("}}", "}").Replace("},", "}");
                    var av = v.Split("}");
                    for (var i = 0; i < av.Length; i++)
                    {
                        if (!string.IsNullOrEmpty(av[i]))
                            av[i] = "new []" + av[i] + "}";
                    }
                    realValue = $"new {realType}{{{string.Join(",", av)}}}";
                }
                else if (realType.Equals("string[]", StringComparison.CurrentCultureIgnoreCase))
                {
                    var array = realValue.Replace("{", "").Replace("}", "").Trim().Split(',');
                    var result = "";
                    foreach (var a in array)
                    {
                        if (!string.IsNullOrEmpty(result))
                            result += ",";
                        if (!a.Trim().StartsWith("\""))
                            result += "\"";
                        result += a;
                        if (!a.Trim().EndsWith("\""))
                            result += "\"";
                    }
                    realValue = $"new {realType}{{{result}}}";
                }
                else
                    realValue = $"new {realType}{{{realValue.Replace("{", "").Replace("}", "")}}}";
            }
            else if (realType == "float")
            {
                if (!string.IsNullOrEmpty(realValue))
                    realValue = $"{realValue}{(!realValue.EndsWith("f") ? "f" : "")}";
            }
            else if (realType.StartsWith("Dictionary<"))
            {
                var rawvalue = realValue;
                if (string.IsNullOrEmpty(rawvalue))
                    return "null";
                var avalues = realValue.Replace("{", "").Replace("}", "").Split(',');
                realValue = "{";
                foreach (var value in avalues)
                {
                    if (string.IsNullOrEmpty(value.Trim()))
                        continue;
                    var avalue = value.Split(new char[] { '=', ':' }, 2);
                    if (avalue.Length != 2)
                    {
                        return TryParse(rawvalue, realType);//解析自定义类
                    }
                    var a0 = avalue[0].Trim();
                    var a1 = avalue[1].Trim();
                    if (realType.Contains("string,"))
                        a0 = a0.StartsWith("\"") ? a0 : $"\"{a0}\"";
                    if (realType.Contains(",string"))
                        a1 = a1.StartsWith("\"") ? a1 : $"\"{a1}\"";
                    if (realType.Contains(",float"))
                        a1 = a1.EndsWith("f") ? a1 : $"{a1}f";
                    realValue += $"{{{a0},{a1}}},";
                }
                realValue += "}";
                realValue = $"new {realType}(){realValue}";
            }
            return realValue;
        }

        private static string TryParse(string value, string realType)
        {
            var res = $"new {realType}{{";
            var svarlue = value.Replace(", ", ",");
            var mapType = realType.Replace(">", "").Split(',')[1];
            var avalue = svarlue.Split(",{");
            for (var i = 0; i < avalue.Length; i++)
            {
                var v = avalue[i].Replace("{", "").Replace("}", "");
                if (i % 2 != 0)
                    res += "new " + mapType + "(){" + v + "}},";
                else
                    res += "{" + v + ",";
            }
            return res + "}";
        }

        private static string GetCellText(ICell cell)
        {
            try
            {
                if (cell.CellType != CellType.Formula)
                    return cell.ToString();
                if (cell.CachedFormulaResultType == CellType.String)
                    return cell.StringCellValue;
                if (cell.CachedFormulaResultType != CellType.Numeric)
                    return cell.StringCellValue;
                return cell.NumericCellValue.ToString();
            }
            catch
            {
                cell.SetCellType(CellType.String);
                return cell.StringCellValue;
            }
        }
    }

}