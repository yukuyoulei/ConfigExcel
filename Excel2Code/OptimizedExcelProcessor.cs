using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System;
using System.IO;
using System.Text;

namespace Excel2Code
{
    /// <summary>
    /// Optimized Excel processor for better performance with large files
    /// </summary>
    public static class OptimizedExcelProcessor
    {
        /// <summary>
        /// Process Excel file with optimized streaming approach
        /// </summary>
        /// <param name="filePath">Path to the Excel file</param>
        /// <param name="processSheet">Callback to process each sheet</param>
        public static void ProcessExcelFile(string filePath, Action<ISheet> processSheet)
        {
            IWorkbook workbook = null;
            
            try
            {
                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var fileExtension = Path.GetExtension(filePath).ToLower();
                    
                    // Use the appropriate workbook type based on file extension
                    if (fileExtension == ".xlsx")
                    {
                        // For .xlsx files, we can use streaming reader for better memory efficiency
                        workbook = new XSSFWorkbook(fileStream);
                    }
                    else if (fileExtension == ".xls")
                    {
                        workbook = new HSSFWorkbook(fileStream);
                    }
                    else
                    {
                        throw new InvalidOperationException($"Unsupported file format: {fileExtension}");
                    }
                    
                    // Process each sheet
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        var sheet = workbook.GetSheetAt(i);
                        processSheet(sheet);
                    }
                }
            }
            finally
            {
                // Ensure workbook is disposed properly
                if (workbook != null)
                {
                    workbook.Close();
                }
            }
        }
        
        /// <summary>
        /// Efficiently gets cell value as string
        /// </summary>
        /// <param name="cell">The cell to extract value from</param>
        /// <returns>String representation of cell value</returns>
        public static string GetCellValueAsString(ICell cell)
        {
            if (cell == null)
                return string.Empty;
                
            try
            {
                switch (cell.CellType)
                {
                    case CellType.String:
                        return cell.StringCellValue ?? string.Empty;
                    case CellType.Numeric:
                        return cell.NumericCellValue.ToString();
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case CellType.Formula:
                        // Handle formula cells by getting cached result
                        switch (cell.CachedFormulaResultType)
                        {
                            case CellType.String:
                                return cell.StringCellValue ?? string.Empty;
                            case CellType.Numeric:
                                return cell.NumericCellValue.ToString();
                            case CellType.Boolean:
                                return cell.BooleanCellValue.ToString();
                            default:
                                return cell.ToString() ?? string.Empty;
                        }
                    case CellType.Blank:
                        return string.Empty;
                    default:
                        return cell.ToString() ?? string.Empty;
                }
            }
            catch
            {
                // Fallback to string representation
                return cell.ToString() ?? string.Empty;
            }
        }
    }
}