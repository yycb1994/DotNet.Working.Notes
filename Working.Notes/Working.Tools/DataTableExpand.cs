using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using Working.Tools.AttributeExpand;



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
                        var attribute = property.GetCustomAttribute<DataTableFieldNameAttribute>();
                        var columnName = attribute?.ColumnName ?? property.Name;

                        if (dataTable.Columns.Contains(columnName))
                        {
                            var value = row[columnName];
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
        /// 将泛型集合类转换成DataTable
        /// </summary>
        /// <typeparam name="T">集合项类型</typeparam>
        /// <param name="list">集合</param>
        /// <param name="tableName">表名</param>
        /// <returns>数据集(表)</returns>
        public static DataTable ToDataTable<T>(this IList<T> list, string tableName = null)
        {
            var result = new DataTable(tableName);

            if (list.Count == 0)
            {
                return result;
            }

            var properties = typeof(T).GetProperties();
            result.Columns.AddRange(properties.Select(p =>
            {
                var columnType = p.PropertyType;
                if (columnType.IsGenericType && columnType.GetGenericTypeDefinition() == typeof(Nullable<>))
                {
                    columnType = Nullable.GetUnderlyingType(columnType);
                }
                return new DataColumn(p.GetCustomAttribute<DataTableFieldNameAttribute>()?.ColumnName ?? p.Name, columnType);
            }).ToArray());

            list.ToList().ForEach(item => result.Rows.Add(properties.Select(p => p.GetValue(item)).ToArray()));

            return result;
        }

        /// <summary>
        /// 给DataTable增加一个自增列
        /// 如果DataTable 存在 identityid 字段  则 直接返回DataTable 不做任何处理
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <returns>返回Datatable 增加字段 identityid </returns>
        public static DataTable AddIdentityColumn(this DataTable dt,string columnName= "identityid")
        {
            if (!dt.Columns.Contains(columnName))
            {
                DataColumn identityColumn = new DataColumn(columnName);
                dt.Columns.Add(identityColumn);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dt.Rows[i][columnName] = (i + 1).ToString();
                }

                dt.Columns[columnName].SetOrdinal(0); // 将列排在第一位
            }

            return dt;
        }

        /// <summary>
        /// 将 DataTable 导入到 Excel 文件
        /// </summary>
        /// <param name="dataTable">要导入的 DataTable</param>
        /// <param name="saveFullPath">保存文件的完整路径</param>
        /// <param name="title">标题内容</param>
        /// <param name="sheetName">工作表名称，默认为 "Sheet1"</param>
        public static void ImportExcel(this DataTable dataTable, string saveFullPath, string title, string sheetName = "Sheet1")
        {
            FileHelper.CreateDirectoryPath(Path.GetDirectoryName(saveFullPath));

            IWorkbook workbook = new XSSFWorkbook(); // 创建 .xlsx 文件
                                                     // IWorkbook workbook = new HSSFWorkbook(); // 创建 .xls 文件

            ISheet sheet = workbook.CreateSheet(sheetName); // 创建工作表，名称为 "Sheet1"


            // 设置标题样式
            var cellStyleFont = ExcelHelper.CreateStyle(workbook, HorizontalAlignment.Center, VerticalAlignment.Center, 20, true, 700, "楷体", true, false, false, true, FillPattern.SolidForeground, HSSFColor.Coral.Index, HSSFColor.White.Index,
                    FontUnderlineType.None, FontSuperScript.None, false);


            //第一行表单
            var row = ExcelHelper.CreateRow(sheet, 0, 28);
            var cell = row.CreateCell(0);
            CellRangeAddress region = new CellRangeAddress(0, 0, 0, 20);
            sheet.AddMergedRegion(region);
            cell.SetCellValue(title);//合并单元格后，只需对第一个位置赋值即可（TODO:顶部标题）
            cell.CellStyle = cellStyleFont;

            //二级标题列样式设置
            var headTopStyle = ExcelHelper.CreateStyle(workbook, HorizontalAlignment.Center, VerticalAlignment.Center, 15, true, 700, "楷体", true, false, false, true, FillPattern.SolidForeground, HSSFColor.Grey25Percent.Index, HSSFColor.Black.Index,
            FontUnderlineType.None, FontSuperScript.None, false);

            row = ExcelHelper.CreateRow(sheet, 1, 28);
            // 写入表头
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {

                cell = ExcelHelper.CreateCells(row, headTopStyle, i, dataTable.Columns[i].ColumnName);
                sheet.SetColumnWidth(i, 5000);
            }


            // 设置数据行样式
            var cellStyle = ExcelHelper.CreateStyle(workbook, HorizontalAlignment.Center, VerticalAlignment.Center, 10, true, 400);


            // 写入数据
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                row = ExcelHelper.CreateRow(sheet, i + 2, 20); //sheet.CreateRow(i+2);//在上面表头的基础上创建行
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    cell = ExcelHelper.CreateCells(row, cellStyle, j, dataTable.Rows[i][j].ToString());
                    //cell.SetCellValue(dataTable.Rows[i][j].ToString());
                    //cell.CellStyle = dataStyle;
                }
            }


            using (FileStream fs = new FileStream(saveFullPath, FileMode.Create))
            {
                workbook.Write(fs);
            }
        }

    }

}
