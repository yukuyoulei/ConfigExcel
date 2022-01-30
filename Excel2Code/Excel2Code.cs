using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Code
{
	public static class Excel2Code
	{
		public static void Generate(string file)
		{
			IWorkbook workbook;
			var finfo = new FileInfo(file);
			string fileExt = finfo.Extension.ToLower();
			var classname = finfo.Name.Substring(0, finfo.Name.Length - fileExt.Length);
			var classes = "";
			var res = "";
			res = Utils.AppendLine(res, "//本文件为自动生成，请勿手动修改");
			res = Utils.AppendLine(res, "using System.Collections.Generic;");
			res = Utils.AppendLine(res, $"public partial class {classname}");
			res = Utils.AppendLine(res, "{");
			using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read))
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
					else if (!name.Contains(" "))
					{
						classes = Utils.AppendLine(classes, GenerateSubDataClass(sheet));
					}
					else if (name.Contains(" "))
					{
						if (sheet.LastRowNum - sheet.FirstRowNum == 1)
							res = Utils.AppendLine(res, GenerateSubData(sheet));
						else
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

		private static string GenerateSubDataList(ISheet sheet)
		{
			var classname = sheet.SheetName.Split(' ')[0];
			var paramname = sheet.SheetName.Split(' ')[1];
			var res = "";
			res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}public static List<{classname}> {paramname} = new List<{classname}>()");
			res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}{{");
			var hrow = sheet.GetRow(0);
			var columns = hrow.LastCellNum;
			for (int i = 1; i <= sheet.LastRowNum; i++)
			{
				res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}new {classname}");
				res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}{{");
				var row = sheet.GetRow(i);
				for (var j = 0; j < columns; j++)
				{
					res = Utils.AppendLine(res, $"{Utils.GetTabs(3)}{hrow.GetCell(j)} = {row.GetCell(j)},");
				}
				res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}}},");
			}
			res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}}};");
			return res;
		}

		private static string GenerateSubDataClass(ISheet sheet)
		{
			var res = "";
			res = Utils.AppendLine(res, $"public class {sheet.SheetName}");
			res = Utils.AppendLine(res, "{");
			for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
			{
				var row = sheet.GetRow(i);
				var valueType = row.GetCell(0);
				var valueName = row.GetCell(1);
				res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}public {valueType} {valueName};");
			}
			res = Utils.AppendLine(res, "}");
			return res;
		}

		private static string GenerateSubData(ISheet sheet)
		{
			var classname = sheet.SheetName.Split(' ')[0];
			var res = "";
			res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}public static {sheet.SheetName} = new {classname}()");
			res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}{{");
			var hrow = sheet.GetRow(0);
			var columns = hrow.LastCellNum;
			for (int i = 1; i <= sheet.LastRowNum; i++)
			{
				var row = sheet.GetRow(i);
				for (var j = 0; j < columns; j++)
				{
					res = Utils.AppendLine(res, $"{Utils.GetTabs(2)}{hrow.GetCell(j)} = {row.GetCell(j)},");
				}
			}
			res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}}};");
			return res;
		}

		/// <summary>
		/// 获取单元格类型
		/// </summary>
		/// <param name="cell">目标单元格</param>
		/// <returns></returns>
		private static string GetValueType(ICell cell)
		{
			if (cell == null)
				throw new Exception($"Missing cell");
			switch (cell.CellType)
			{
				case CellType.Boolean:
					return "bool";
				case CellType.Numeric:
					return "int";
				case CellType.String:
				case CellType.Blank:
				case CellType.Error:
				case CellType.Formula:
				default:
					return "string";
			}
		}
		private static string GenerateDesc(ISheet sheet)
		{
			var res = "";
			for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
			{
				var row = sheet.GetRow(i);
				var valueType = row.GetCell(0);
				var valueName = row.GetCell(1);
				var value = row.GetCell(2);
				var realType = valueType == null ? GetValueType(value) : valueType.ToString();
				var realValue = value.ToString();
				if (realType == "bool")
					realValue = realValue.ToLower();
				res = Utils.AppendLine(res, $"{Utils.GetTabs(1)}public static {realType} {valueName} = {realValue};");
			}
			return res;
		}
	}

}
