using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Excel2Code
{
    public static class Excel2Code
    {
        public static string ExportingFile;
        public static string ExportingSheet;
        public static string ExportingColumn;
        private const string historyFile = "history.txt";
        private const string ignoreHistoryFile = "ignoreHistory.txt";
        private static Lazy<List<string>> ignoreHistoryFiles = new Lazy<List<string>>(() =>
        {
            if (!File.Exists(ignoreHistoryFile))
                return new();
            var all = File.ReadAllLines(ignoreHistoryFile);
            return all.ToList().Where(r => !string.IsNullOrEmpty(r.Trim())).ToList();
        });
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
                    Console.WriteLine($"{ignoreHistoryFiles.Value.Count} files cannot be jumped");
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
                    var fname = finfo.Name.Split('.')[0].Trim();
                    if (ignores.Contains(fname))
                        continue;
                    if (fileExt == "xls" || fileExt == ".xlsx")
                    {
                        var curt = (finfo.LastWriteTime - new DateTime(1970, 1, 1)).Ticks;
                        if (excelHistories.TryGetValue(fname, out var t) && curt == t)
                        {
                            if (!ignoreHistoryFiles.Value.Contains(fname))
                            {
                                Console.WriteLine($"Already generated {f}");
                                continue;
                            }
                        }
                        Excel2Code.Generate(f, outdir);
                        excelHistories[fname] = curt;
                    }
                }
            }
            if (!string.IsNullOrEmpty(compiledir))
            {
                //获取目标目录中的所有C#文件
                List<string> cSharpFiles = Directory.GetFiles(compiledir, "*.cs").Where(r => !r.Contains("Configer")).ToList();
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
                var dllpath = Path.Combine(dir, "temp.dll");
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
                File.Delete(dllpath);
            }
        }

        private static List<string> subClasses = new List<string>();
        static List<ISheet> pendingSheets = new();
        private static string configClassName;
        private static HashSet<string> customNamespaces = new HashSet<string>();

        public static void Generate(string file, string outdir)
        {
            Console.WriteLine($"Generated {file}");
            ExportingFile = file;

            // Clear custom namespaces for this file
            customNamespaces.Clear();

            var finfo = new FileInfo(file);
            string fileExt = finfo.Extension.ToLower();
            configClassName = finfo.Name.Substring(0, finfo.Name.Length - fileExt.Length);
            if (!configClassName.Contains("Config_"))
                configClassName = $"Config_{configClassName}";
            var classes = new StringBuilder();
            var res = new StringBuilder();
            res.AppendLine("//本文件为自动生成，请勿手动修改");
            res.AppendLine("//");
            res.AppendLine("using System;");
            res.AppendLine("using System.Collections.Generic;");

            // First pass: collect all ##namespace directives by scanning each sheet
            OptimizedExcelProcessor.ProcessExcelFile(file, sheet =>
            {
                // Collect ##namespace directives from all sheets
                // This will be done by GetDataStartRow() internally
                var name = sheet.SheetName;
                if (!name.StartsWith("#") && !name.StartsWith("Sheet"))
                {
                    GetDataStartRow(sheet); // This collects ##namespace directives
                }
            });

            // Add custom namespace using statements
            foreach (var ns in customNamespaces.OrderBy(x => x))
            {
                res.AppendLine($"using {ns};");
            }

            // Start class definition
            //res.AppendLine("namespace ConfigExcel");
            //res.AppendLine("{");
            res.AppendLine($"public partial class {configClassName}");
            res.AppendLine("{");

            // Second pass: generate actual code
            OptimizedExcelProcessor.ProcessExcelFile(file, sheet =>
            {
                var name = sheet.SheetName;
                ExportingSheet = name;

                if (name.StartsWith("#")) // 参数表
                {
                    res.AppendLine(GenerateDesc(sheet));
                }
                else if (name.StartsWith("Sheet"))
                    return;
                else
                {
                    if (sheet.SheetName.StartsWith("="))
                    {
                        pendingSheets.Add(sheet);
                        return;
                    }
                    var subclass = GenerateSubDataClass(sheet);
                    if (!string.IsNullOrEmpty(subclass))
                        classes.AppendLine(subclass);
                    var headcount = sheet.LastRowNum - sheet.FirstRowNum;
                    //if (headcount == 3)
                    //    res.AppendLine(GenerateSubData(sheet));
                    //else if (headcount > 3)
                    res.AppendLine(GenerateSubDataList(sheet));
                    pendingSheets.Clear();
                }
            });

            res.AppendLine("}");
            res.AppendLine(classes.ToString());
            //res.AppendLine("}");

            var path = $"{outdir}/{configClassName}.cs";
            File.WriteAllText(path, res.ToString().Replace("\r\n", "----").Replace("\n", "\r\n").Replace("----", "\r\n"));
            Console.WriteLine($"Generated {path}");
        }

        /// <summary>
        /// Process preprocessor directives (##namespace, etc.) from the first row of a sheet
        /// </summary>
        private static void ProcessPreprocessorDirectives(ISheet sheet)
        {
            if (sheet == null || sheet.LastRowNum < 0)
                return;

            var firstRow = sheet.GetRow(sheet.FirstRowNum);
            if (firstRow == null)
                return;

            var firstCell = firstRow.GetCell(0);
            if (firstCell == null)
                return;

            var cellValue = firstCell.ToString()?.Trim();
            if (string.IsNullOrEmpty(cellValue))
                return;

            // Check for ## preprocessor directives
            if (cellValue.StartsWith("##"))
            {
                var directive = cellValue.Substring(2).Trim();

                // ##namespace directive: ##namespace Server, Client
                if (directive.StartsWith("namespace"))
                {
                    var namespaceDecl = directive.Substring("namespace".Length).Trim();
                    if (!string.IsNullOrEmpty(namespaceDecl))
                    {
                        // Split by comma and add each namespace
                        var namespaces = namespaceDecl.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var ns in namespaces)
                        {
                            var trimmedNs = ns.Trim();
                            if (!string.IsNullOrEmpty(trimmedNs))
                            {
                                customNamespaces.Add(trimmedNs);
                            }
                        }
                    }
                }
                // Future directives can be added here (##import, ##define, etc.)
            }
        }

        /// <summary>
        /// Check if a row should be skipped. Returns true if the row starts with # or ##.
        /// If it starts with ##, process the preprocessor directive first.
        /// </summary>
        private static bool ShouldSkipRow(ICell firstCell)
        {
            if (firstCell == null)
                return false;

            var cellValue = firstCell.ToString()?.Trim();
            if (string.IsNullOrEmpty(cellValue))
                return false;

            // Check for ## preprocessor directives
            if (cellValue.StartsWith("##"))
            {
                // Process the directive (for data rows, not sheet-level)
                var directive = cellValue.Substring(2).Trim();

                // ##namespace directive in data rows
                if (directive.StartsWith("namespace"))
                {
                    var namespaceDecl = directive.Substring("namespace".Length).Trim();
                    if (!string.IsNullOrEmpty(namespaceDecl))
                    {
                        var namespaces = namespaceDecl.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var ns in namespaces)
                        {
                            var trimmedNs = ns.Trim();
                            if (!string.IsNullOrEmpty(trimmedNs))
                            {
                                customNamespaces.Add(trimmedNs);
                            }
                        }
                    }
                }
                // Skip this row after processing
                return true;
            }

            // Check for single # comment
            if (cellValue.StartsWith("#"))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Get the actual data start row index, skipping all ## preprocessor directives and # comments at the beginning
        /// Note: User guarantees that all ##namespace directives are at the top (before any data rows)
        /// Single # comment rows can appear anywhere and will be skipped
        /// </summary>
        private static int GetDataStartRow(ISheet sheet)
        {
            int startRow = 0;

            // Skip all ## preprocessor directives and # comments at the beginning
            while (startRow <= sheet.LastRowNum)
            {
                var row = sheet.GetRow(startRow);
                if (row == null)
                {
                    startRow++;
                    continue;
                }

                var firstCell = row.GetCell(0);
                if (firstCell == null)
                {
                    break;
                }

                var cellValue = firstCell.ToString()?.Trim();
                if (string.IsNullOrEmpty(cellValue))
                {
                    break;
                }

                // Process ## directives (they are guaranteed to be at the top)
                if (cellValue.StartsWith("##"))
                {
                    // Process the directive
                    var directive = cellValue.Substring(2).Trim();
                    if (directive.StartsWith("namespace"))
                    {
                        var namespaceDecl = directive.Substring("namespace".Length).Trim();
                        if (!string.IsNullOrEmpty(namespaceDecl))
                        {
                            var namespaces = namespaceDecl.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var ns in namespaces)
                            {
                                var trimmedNs = ns.Trim();
                                if (!string.IsNullOrEmpty(trimmedNs))
                                {
                                    customNamespaces.Add(trimmedNs);
                                }
                            }
                        }
                    }
                    // Skip this row after processing
                    startRow++;
                    continue;
                }

                // Skip # comment rows
                if (cellValue.StartsWith("#"))
                {
                    startRow++;
                    continue;
                }

                // Found the first non-# and non-## row (actual data header)
                // This is the start of data section
                break;
            }

            return startRow;
        }

        /// <summary>
        /// Get the next non-comment row starting from the given row index
        /// Skips rows that start with # or ##
        /// Returns tuple of (IRow, rowIndex)
        /// </summary>
        private static (IRow row, int index) GetNextNonCommentRow(ISheet sheet, int startFromRow)
        {
            for (int i = startFromRow; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                    continue;

                var firstCell = row.GetCell(0);
                if (firstCell == null)
                    return (row, i);

                var cellValue = firstCell.ToString()?.Trim();
                if (string.IsNullOrEmpty(cellValue))
                    return (row, i);

                // Skip comment rows
                if (cellValue.StartsWith("#"))
                    continue;

                // Found a non-comment row
                return (row, i);
            }

            return (null, -1);
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
            var res = new StringBuilder();
            res.AppendLine($"{Utils.GetTabs(1)}public partial class {classname}");
            res.AppendLine($"{Utils.GetTabs(1)}{{");

            // Get actual start row, skipping ## directives and # comments
            var startRow = GetDataStartRow(sheet);
            var (rowComment, commentIdx) = GetNextNonCommentRow(sheet, startRow);
            var (rowType, typeIdx) = GetNextNonCommentRow(sheet, commentIdx + 1);
            var (rowParamName, paramIdx) = GetNextNonCommentRow(sheet, typeIdx + 1);
            for (var i = 0; i < rowType.LastCellNum; i++)
            {
                var paramName = rowParamName.GetCell(i)?.ToString();
                if (string.IsNullOrEmpty(paramName))
                    break;
                var pn = GetRealType(rowParamName.GetCell(i));
                if (pn == "ignore")
                    continue;
                var realtype = GetRealType(rowType.GetCell(i));
                if (realtype == "ignore")
                    continue;
                res.AppendLine($"{Utils.GetTabs(2)}public {realtype} {paramName}; /*{rowComment.GetCell(i)}*/");
            }
            res.AppendLine($"{Utils.GetTabs(1)}}}");
            return res.ToString();
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
            var res = new StringBuilder();

            // Get actual start row, skipping ## directives and # comments
            var startRow = GetDataStartRow(sheet);
            var (rowComment, commentIdx) = GetNextNonCommentRow(sheet, startRow);
            var (rowType, typeIdx) = GetNextNonCommentRow(sheet, commentIdx + 1);
            var (paramNameRow, paramIdx) = GetNextNonCommentRow(sheet, typeIdx + 1);
            var columns = paramNameRow.LastCellNum;
            var firstRowType = GetRealType(rowType.GetCell(0));
            var genericList = rowComment.GetCell(0).ToString().ToLower().Contains("list");
            var allres = new StringBuilder();
            var dataStartRow = paramIdx + 1; // Data starts after param name row
            if (genericList)
            {
                res.AppendLine($"{Utils.GetTabs(2)}public static List<{classname}> {paramname} = new List<{classname}>()");
                res.AppendLine($"{Utils.GetTabs(2)}{{");
                for (int i = dataStartRow; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    // Skip rows that start with # or ##
                    var firstCell = row?.GetCell(0);
                    if (ShouldSkipRow(firstCell))
                        continue;

                    res.AppendLine($"{Utils.GetTabs(3)}new {classname}");
                    res.AppendLine($"{Utils.GetTabs(3)}{{");
                    for (var j = 0; j < columns; j++)
                    {
                        var realValue = CellToString(row.GetCell(j), rowType.GetCell(j).ToString());
                        if (string.IsNullOrEmpty(realValue))
                            continue;
                        var paramName = paramNameRow.GetCell(j).ToString();
                        if (string.IsNullOrEmpty(paramName))
                            break;
                        res.AppendLine($"{Utils.GetTabs(4)}{paramName} = {realValue},");
                    }
                    res.AppendLine($"{Utils.GetTabs(3)}}},");
                }

            }
            else
            {
                if (asheetname.Length < 2)
                    paramname = $"d{classname}";
                var keyType = rowType.GetCell(0).ToString().Split(':')[0];
                var keyParamName = paramNameRow.GetCell(0).ToString();
                var dicres = new StringBuilder();
                dicres.AppendLine($"{Utils.GetTabs(2)}public static Dictionary<{keyType}, {classname}> {paramname} = new Dictionary<{keyType}, {classname}>();");
                dicres.AppendLine($"{Utils.GetTabs(2)}public static {classname} OnGetFrom_{paramname}({keyType} id)");
                dicres.AppendLine($"{Utils.GetTabs(2)}{{");
                dicres.AppendLine($"{Utils.GetTabs(3)}{classname} data = null;");
                dicres.AppendLine($"{Utils.GetTabs(3)}if (d{classname}.TryGetValue(id, out data))");
                dicres.AppendLine($"{Utils.GetTabs(3)}{{");
                dicres.AppendLine($"{Utils.GetTabs(4)}return data;");
                dicres.AppendLine($"{Utils.GetTabs(3)}}}");

                dicres.AppendLine($"{Utils.GetTabs(3)}var t = typeof({configClassName});");
                if (firstRowType == "string")
                {
                    dicres.AppendLine($"{Utils.GetTabs(3)}var idx = all{classname}s.IndexOf(id);");
                    dicres.AppendLine($"{Utils.GetTabs(3)}if (idx == -1) return null;");
                    dicres.AppendLine($"{Utils.GetTabs(3)}var m = t.GetMethod($\"Create{classname}_{{idx}}\", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);");
                }
                else
                    dicres.AppendLine($"{Utils.GetTabs(3)}var m = t.GetMethod($\"Create{classname}_{{id}}\", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);");
                dicres.AppendLine($"{Utils.GetTabs(3)}if (m == null) return null;");
                dicres.AppendLine($"{Utils.GetTabs(3)}data = m.Invoke(null, null) as {classname};");
                dicres.AppendLine($"{Utils.GetTabs(3)}d{classname}[id] = data;");
                dicres.AppendLine($"{Utils.GetTabs(3)}return data;");
                dicres.AppendLine($"{Utils.GetTabs(2)}}}");

                //dicres.AppendLine($"{Utils.GetTabs(3)}if (s_{classname}Initializers.TryGetValue(id, out var initializer))");
                //dicres.AppendLine($"{Utils.GetTabs(3)}{{");
                //dicres.AppendLine($"{Utils.GetTabs(4)}data = initializer();");
                //dicres.AppendLine($"{Utils.GetTabs(4)}d{classname}[id] = data;");
                //dicres.AppendLine($"{Utils.GetTabs(4)}return data;");
                //dicres.AppendLine($"{Utils.GetTabs(3)}}}");
                //dicres.AppendLine($"{Utils.GetTabs(3)}return null;");
                //dicres.AppendLine($"{Utils.GetTabs(2)}}}");

                //dicres.AppendLine($"{Utils.GetTabs(2)}private static readonly Dictionary<{keyType}, Func<{classname}>> s_{classname}Initializers = new()");
                //dicres.AppendLine($"{Utils.GetTabs(2)}{{");
                //for (int i = 3; i <= sheet.LastRowNum; i++)
                //{
                //    var row = sheet.GetRow(i);
                //    if (row == null)
                //        continue;
                //    var r0 = row.GetCell(0);
                //    if (r0 == null)
                //        break;
                //    var srow0 = r0.ToString();
                //    if (string.IsNullOrEmpty(srow0))
                //        break;
                //    var rawrow0 = srow0;
                //    if (firstRowType == "string")
                //    {
                //        srow0 = $"\"{srow0}\"";
                //        rawrow0 = (i - 3).ToString();
                //    }
                //    dicres.AppendLine($"{Utils.GetTabs(3)}{{{srow0}, Create{classname}_{rawrow0}}},");
                //}
                //dicres.AppendLine($"{Utils.GetTabs(2)}}};");

                for (int i = dataStartRow; i <= sheet.LastRowNum; i++)
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
                    // Skip rows that start with # or ##
                    if (ShouldSkipRow(r0))
                        continue;
                    var rawrow0 = srow0;
                    if (firstRowType == "string")
                    {
                        srow0 = $"\"{srow0}\"";
                        rawrow0 = (i - 3).ToString();
                    }
                    dicres.AppendLine($"{Utils.GetTabs(2)}private static {classname} Create{classname}_{rawrow0}()");
                    dicres.AppendLine($"{Utils.GetTabs(2)}{{");
                    dicres.AppendLine($"{Utils.GetTabs(3)}return new {classname}()");
                    dicres.AppendLine($"{Utils.GetTabs(3)}{{");
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
                        if (paramName == "ignore")
                            continue;
                        dicres.AppendLine($"{Utils.GetTabs(4)}{paramName} = {realValue},");
                    }
                    dicres.AppendLine($"{Utils.GetTabs(3)}}};");
                    dicres.AppendLine($"{Utils.GetTabs(2)}}}");
                }

                //dicres.AppendLine($"{Utils.GetTabs(3)}switch ({keyParamName})");
                //dicres.AppendLine($"{Utils.GetTabs(3)}{{");
                var datares = new StringBuilder();
                SheetToString(datares, firstRowType, ref allres, columns, paramname
                    , classname, rowType, paramNameRow
                    , sheet, dataStartRow);
                foreach (var s in pendingSheets)
                {
                    var pendingStartRow = GetDataStartRow(s);
                    var (pendingRowComment, pendingCommentIdx) = GetNextNonCommentRow(s, pendingStartRow);
                    var (pendingRowType, pendingTypeIdx) = GetNextNonCommentRow(s, pendingCommentIdx + 1);
                    var (pendingParamNameRow, pendingParamIdx) = GetNextNonCommentRow(s, pendingTypeIdx + 1);
                    var pendingDataStartRow = pendingParamIdx + 1;
                    SheetToString(datares, firstRowType, ref allres, columns, paramname
                        , classname, rowType, paramNameRow
                        , s, pendingDataStartRow);
                }
                //if (datares.Length == 0)
                //return res.ToString();
                res.Append(dicres.ToString());
                //res.Append(datares.ToString());
                //res.AppendLine($"{Utils.GetTabs(3)}}}");
                //res.AppendLine($"{Utils.GetTabs(3)}return null;");
            }
            //res.AppendLine($"{Utils.GetTabs(2)}}}{(genericList ? ";" : "")}");
            if (allres.Length > 0)
            {
                res.AppendLine($"{Utils.GetTabs(2)}public static List<{firstRowType}> all{classname}s = new List<{firstRowType}>()");
                res.AppendLine($"{Utils.GetTabs(3)}{{");
                res.AppendLine(allres.ToString());
                res.AppendLine($"{Utils.GetTabs(3)}}};");
                res.AppendLine($"{Utils.GetTabs(2)}public static {classname} OnGetFrom_{classname}ByIndex(int index){{if (index < 0 || index >= all{classname}s.Count) return null; return OnGetFrom_{paramname}(all{classname}s[index]);}}");
            }
            return res.ToString();
        }

        private static void SheetToString(StringBuilder res, string firstRowType, ref StringBuilder allres
            , short columns, string paramname, string classname
            , IRow? rowType, IRow? paramNameRow
            , ISheet sheet, int dataStartRow)
        {
            for (int i = dataStartRow; i <= sheet.LastRowNum; i++)
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
                // Skip rows that start with # or ##
                if (ShouldSkipRow(r0))
                    continue;
                if (firstRowType == "string")
                    srow0 = $"@\"{srow0}\"";
                allres.AppendLine($"{Utils.GetTabs(4)}{srow0},");
                res.AppendLine($"{Utils.GetTabs(4)}case {srow0}:");
                res.AppendLine($"{Utils.GetTabs(5)}if (!{paramname}.TryGetValue({srow0}, out data))");
                res.AppendLine($"{Utils.GetTabs(5)}{{");
                res.AppendLine($"{Utils.GetTabs(6)}data = new {classname}()");
                res.AppendLine($"{Utils.GetTabs(6)}{{");
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
                    if (paramName == "ignore")
                        continue;
                    res.AppendLine($"{Utils.GetTabs(7)}{paramName} = {realValue},");
                }
                res.AppendLine($"{Utils.GetTabs(6)}}};");
                res.AppendLine($"{Utils.GetTabs(6)}{paramname}[{srow0}] = data;");
                res.AppendLine($"{Utils.GetTabs(5)}}}");
                res.AppendLine($"{Utils.GetTabs(5)}return {paramname}[{srow0}];");
            }
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
            var res = new StringBuilder();
            res.AppendLine($"{Utils.GetTabs(2)}public static {classname} {paramname} = new {classname}()");
            res.AppendLine($"{Utils.GetTabs(2)}{{");
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
                    res.AppendLine($"{Utils.GetTabs(3)}{paramName} = {realValue},");
                }
            }
            res.AppendLine($"{Utils.GetTabs(2)}}};");
            return res.ToString();
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
            var res = new StringBuilder();
            for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                    continue;

                var valueType = row.GetCell(0);
                // Skip rows that start with # or ##
                if (ShouldSkipRow(valueType))
                    continue;

                var valueNameCell = row.GetCell(1);
                if (valueNameCell == null)
                    return res.ToString();

                var valueName = valueNameCell.ToString();
                var value = row.GetCell(2);
                var comment = row.GetCell(3);
                var realType = valueType == null ? GetValueType(value, true) : GetRealType(valueType);
                var realValue = CellToString(value, realType);
                if (realValue == null)
                    return res.ToString();
                if (string.IsNullOrEmpty(valueName))
                    return res.ToString();
                var key = "static";
                if (realType.Equals("string") || realType.Equals("int") || realType.Equals("float") || realType.Equals("long"))
                    key = "const";
                res.AppendLine(
                    $"{Utils.GetTabs(2)}public {key} {realType} {valueName} = {realValue};{(comment == null ? "" : $"/*{comment.ToString()}*/")}" // added
                    );
            }
            return res.ToString();
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
            if (cell == null)
                return "";
            realType = string.IsNullOrEmpty(realType) ? GetValueType(cell, false) : realType;
            realType = realType.Split(':')[0];
            realType = realType.Replace(" ", "");//过滤掉所有空格

            // Handle Enum types BEFORE drepalace: EnumXXX or XXXEnum -> XXX
            // This must be done first to avoid incorrect replacements
            if (realType.StartsWith("Enum") && realType.Length > 4)
            {
                realType = realType.Substring(4); // Remove "Enum" prefix
            }
            else if (realType.EndsWith("Enum") && realType.Length > 4)
            {
                realType = realType.Substring(0, realType.Length - 4); // Remove "Enum" suffix
            }

            // Apply standard replacements (v2, v3, map, etc.)
            foreach (var kv in drepalace)
                realType = realType.Replace(kv.Key, kv.Value);

            return realType;
        }
        private static string CellToString(ICell cell, string realType)
        {
            var realValue = OptimizedExcelProcessor.GetCellValueAsString(cell);
            if (string.IsNullOrEmpty(realValue.Trim()))
                return "";
            realType = GetRealType(cell, realType);

            // Handle Enum values: if value contains a dot and matches the type name, keep it as-is
            // e.g., HeroType.Magic -> HeroType.Magic
            if (realValue.Contains("."))
            {
                var parts = realValue.Split('.');
                if (parts.Length == 2 && parts[0].Trim() == realType)
                {
                    // This is an enum value, return as-is
                    return realValue.Trim();
                }
            }

            if (realType == "bool")
                realValue = realValue.ToLower() == "true" ? "true" : "false";
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
                else if (!string.IsNullOrEmpty(realValue) && realType.StartsWith("string"))
                {
                    var ar = realValue.Split(',');
                    var str = "";
                    foreach (var item in ar)
                    {
                        if (string.IsNullOrEmpty(item))
                            continue;
                        if (!string.IsNullOrEmpty(str))
                            str += ",";
                        if (!item.StartsWith("\""))
                        {
                            str += "\"";
                            str += item;
                        }
                        else
                        {
                            str += item;
                        }
                        if (!item.EndsWith("\""))
                        {
                            str += "\"";
                        }
                    }
                    realValue = str;
                }
                // Handle Enum arrays: keep enum values as-is (e.g., HeroType.Magic,HeroType.Physical)
                // Enum values in arrays should already contain the type name prefix
                // No special processing needed, just wrap in array syntax

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
    }

}
