using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Working.Tools
{
    public class ExcelHelper
    {
        #region 导入 Excel 文件并返回 DataTable 对象。
        /// <summary>
        /// 导入 Excel 文件并返回 DataTable 对象。
        /// </summary>       
        /// <param name="filePath">Excel 文件路径。</param>
        /// <param name="columnValidators">要应用的列值验证器的字典。</param>
        /// <returns>导入的数据表。</returns>
        /// <remarks>
        ///调用示例 :
        ///     var columnValidators = new Dictionary<string, Func<object, bool>>
        ///     {
        ///         { "Column1", value => Convert.ToInt32(value) < 10 },
        ///         { "Column2", value => Convert.ToDecimal(value) > 0 },
        ///          // 添加更多的列和验证函数
        ///      };
        ///     var dataTable = ExcelImporter.ImportExcel(filePath, columnValidators);
        /// </remarks>
        public static DataTable ImportExcel(string filePath, Dictionary<string, Func<object, bool>> columnValidators = null)
        {
            if (!File.Exists(filePath))
            {
                throw new Exception("文件不存在！");
            }

            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                string fileExtension = Path.GetExtension(filePath);
                IWorkbook workbook = null;
                if (fileExtension == ".xlsx") // 2007 version
                {
                    workbook = new XSSFWorkbook(fileStream);
                }
                else if (fileExtension == ".xls") // 2003 version
                {
                    workbook = new HSSFWorkbook(fileStream);
                }
                else
                {
                    throw new Exception($"文件打开失败，文件格式不正确！");
                }
                var worksheet = workbook.GetSheetAt(0);

                var dataTable = new DataTable();

                // 读取表头
                var headerRow = worksheet.GetRow(0);
                for (int col = 0; col < headerRow.LastCellNum; col++)
                {
                    var columnHeader = headerRow.GetCell(col)?.ToString();
                    dataTable.Columns.Add(columnHeader);
                }

                // 读取数据行
                for (int row = 1; row <= worksheet.LastRowNum; row++)
                {
                    var dataRow = dataTable.NewRow();
                    var currentRow = worksheet.GetRow(row);

                    for (int col = 0; col < currentRow.LastCellNum; col++)
                    {
                        var cell = currentRow.GetCell(col);
                        var cellValue = GetCellValue(cell);

                        dataRow[col] = cellValue;
                    }

                    if (columnValidators!=null)
                    {
                        // 验证特定列的值
                        foreach (var columnValidator in columnValidators)
                        {
                            var columnName = columnValidator.Key;
                            var validator = columnValidator.Value;

                            var columnValue = dataRow[columnName];
                            if (columnValue != null && columnValue != DBNull.Value)
                            {
                                if (!validator(columnValue))
                                {
                                    throw new Exception($"第 {row + 1} 行的 {columnName} 列的值不符合要求。");
                                }
                            }
                        }
                    }                  

                    dataTable.Rows.Add(dataRow);
                }

                return dataTable;
            }
        }

        private static object GetCellValue(ICell cell)
        {
            if (cell == null)
                return DBNull.Value;

            switch (cell.CellType)
            {
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                        return cell.DateCellValue;
                    else
                        return cell.NumericCellValue;

                case CellType.String:
                    return cell.StringCellValue;

                case CellType.Boolean:
                    return cell.BooleanCellValue;

                case CellType.Formula:
                    return cell.CellFormula;

                default:
                    return DBNull.Value;
            }
        }
        #endregion
    }
}
