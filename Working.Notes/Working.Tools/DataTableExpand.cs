using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;


namespace Working.Tools
{
    /// <summary>
    /// DataTable 拓展
    /// </summary>
    public static class DataTableExtensions
    {
        /// <summary>
        /// 判断 DataTable 是否为空
        /// </summary>
        /// <param name="dt">要判断的 DataTable</param>
        /// <returns>如果 DataTable 为空，则返回 true；否则返回 false。</returns>
        public static bool TableIsNull(this DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0)
                return true;
            return false;
        }

        /// <summary>
        /// 获取指定列索引的 DataRow 值
        /// </summary>
        /// <param name="row">要获取值的 DataRow</param>
        /// <param name="columnIndex">列索引</param>
        /// <returns>列值的字符串表示，如果为空则返回 null。</returns>
        public static string GetValue(this DataRow row, int columnIndex)
        {
            try
            {
                if (row[columnIndex] == null)
                {
                    return null;
                }
                return row[columnIndex].ToString();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 获取指定列名称的 DataRow 值
        /// </summary>
        /// <param name="row">要获取值的 DataRow</param>
        /// <param name="columnName">列名称</param>
        /// <returns>列值的字符串表示，如果为空则返回 null。</returns>
        public static string GetValue(this DataRow row, string columnName)
        {
            try
            {
                if (row[columnName] == null)
                {
                    return null;
                }
                return row[columnName].ToString();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 获取 DataTable 的第一行第一列的值
        /// </summary>
        /// <param name="dt">要获取值的 DataTable</param>
        /// <returns>第一行第一列的值的字符串表示，如果 DataTable 为空则返回 null。</returns>
        public static string FirstOrDefault(this DataTable dt)
        {
            if (dt.TableIsNull())
            {
                return dt.Rows[0].GetValue(0);
            }
            return null;
        }

        /// <summary>
        /// 将 DataTable 转换为指定类型的对象列表
        /// </summary>
        /// <typeparam name="TModel">目标对象类型</typeparam>
        /// <param name="dataTable">要转换的 DataTable</param>
        /// <returns>转换后的对象列表</returns>
        public static List<TModel> ToObjectList<TModel>(this DataTable dataTable) where TModel : class, new()
        {
            var objectList = new List<TModel>();
            try
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    TModel obj = Activator.CreateInstance<TModel>();

                    foreach (var property in typeof(TModel).GetProperties())
                    {
                        if (dataTable.Columns.Contains(property.Name))
                        {
                            var value = row[property.Name];
                            if (value != DBNull.Value)
                            {
                                property.SetValue(obj, Convert.ChangeType(value, property.PropertyType));
                            }
                        }
                    }
                    objectList.Add(obj);
                }
                return objectList;
            }
            catch (Exception ex)
            {
                throw new Exception("DataTable转换失败:", ex);
            }
        }


        /// <summary>
        /// 将 DataTable 导入到 Excel 文件
        /// </summary>
        /// <param name="dataTable">要导入的 DataTable</param>
        /// <param name="saveFullPath">保存文件的完整路径</param>
        /// <param name="sheetName">工作表名称，默认为 "Sheet1"</param>
        public static void ImportExcel(this DataTable dataTable, string saveFullPath, string sheetName = "Sheet1")
        {
            IWorkbook workbook = new XSSFWorkbook(); // 创建 .xlsx 文件
                                                     // IWorkbook workbook = new HSSFWorkbook(); // 创建 .xls 文件

            ISheet sheet = workbook.CreateSheet(sheetName); // 创建工作表，名称为 "Sheet1"

            // 写入表头
            IRow headerRow = sheet.CreateRow(0);
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                headerRow.CreateCell(i).SetCellValue(dataTable.Columns[i].ColumnName);
            }

            // 写入数据
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                IRow dataRow = sheet.CreateRow(i + 1);
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    dataRow.CreateCell(j).SetCellValue(dataTable.Rows[i][j].ToString());
                }
            }

            using (FileStream fs = new FileStream(saveFullPath, FileMode.Create))
            {
                workbook.Write(fs);
            }
        }

    }

}
