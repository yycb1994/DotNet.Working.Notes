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
    public static class ExcelExpand
    {
        /// <summary>
        /// 打开 Excel 文件并返回工作簿对象。
        /// </summary>
        /// <param name="filePath">Excel 文件的路径。</param>
        /// <returns>打开的 Excel 工作簿对象。</returns>
        public static IWorkbook OpenExcel(this string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("The file does not exist!", filePath);
            IWorkbook workbook;
            // 打开一个文件流来读取 Excel 文件
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // 根据文件扩展名决定使用哪个类来处理 Excel 文件
                if (Path.GetExtension(filePath).ToLower() == ".xlsx")
                    workbook = new XSSFWorkbook(file); // 使用 XSSFWorkbook 处理 .xlsx 文件
                else if (Path.GetExtension(filePath).ToLower() == ".xls")
                    workbook = new HSSFWorkbook(file); // 使用 HSSFWorkbook 处理 .xls 文件
                else
                    throw new Exception("The file is not in Excel format!"); // 抛出异常，文件不是 Excel 格式
            }
            return workbook;
        }

        /// <summary>
        /// 将工作簿保存为字节数组。
        /// </summary>
        /// <param name="workBook">要保存的工作簿对象。</param>
        /// <returns>保存的工作簿字节数组。</returns>
        public static byte[] SaveExcel(this IWorkbook workBook)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                workBook.Write(memoryStream); // 将工作簿写入内存流
                var fileBytes = memoryStream.ToArray(); // 将内存流转换为字节数组
                                                        //File.WriteAllBytes(excelOutPut, fileBytes);
                return fileBytes;
            }
        }

        /// <summary>
        /// 从Excel文件中导入数据到DataTable列表。
        /// </summary>
        /// <param name="filePath">Excel文件路径。</param>
        /// <param name="columnValidators">列验证器字典，用于验证特定列的值。</param>
        /// <returns>包含导入数据的DataTable列表。</returns>
        public static List<DataTable> ImportExcel(this string filePath, Dictionary<string, Func<object, bool>>? columnValidators = null)
        {
            List<DataTable> dataTables = new List<DataTable>();

            var workBook = OpenExcel(filePath);
            var sheetCount = workBook.NumberOfSheets;
            for (int i = 0; i < sheetCount; i++)
            {
                var worksheet = workBook.GetSheetAt(i);
                var dataTable = new DataTable(worksheet.SheetName);

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

                    if (columnValidators != null)
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

                dataTables.Add(dataTable);
            }

            return dataTables;
        }

        /// <summary>
        /// 获取单元格的值。
        /// </summary>
        /// <param name="cell">要获取值的单元格。</param>
        /// <returns>单元格的值。</returns>
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

        private static void CopyRow(IRow sourceRow, IRow targetRow)
        {
            // 确保源行和目标行都不为空
            if ((sourceRow != null) && (targetRow != null))
            {
                // 设置目标行的高度与源行相同
                targetRow.Height = sourceRow.Height;
                // 遍历源行的所有单元格
                for (int i = 0; i < sourceRow.LastCellNum; i++)
                {
                    // 获取源单元格
                    ICell sourceCell = sourceRow.GetCell(i);
                    // 创建目标单元格
                    ICell targetCell = targetRow.CreateCell(i);

                    // 如果源单元格不为空，则复制内容和样式到目标单元格
                    if (sourceCell != null)
                    {
                        // 复制样式
                        targetCell.CellStyle = sourceCell.CellStyle;

                        // 复制注释（如果有）
                        if (sourceCell.CellComment != null)
                        {
                            targetCell.CellComment = sourceCell.CellComment;
                        }

                        // 复制超链接（如果有）
                        if (sourceCell.Hyperlink != null)
                        {
                            targetCell.Hyperlink = sourceCell.Hyperlink;
                        }                     
                        // 根据单元格类型复制值
                        switch (sourceCell.CellType)
                        {
                            case CellType.Blank:
                                targetCell.SetCellValue(sourceCell.StringCellValue);
                                break;
                            case CellType.Boolean:
                                targetCell.SetCellValue(sourceCell.BooleanCellValue);
                                break;
                            case CellType.Error:
                                targetCell.SetCellErrorValue(sourceCell.ErrorCellValue);
                                break;
                            case CellType.Formula:
                                targetCell.SetCellFormula(sourceCell.CellFormula);
                                break;
                            case CellType.Numeric:
                                targetCell.SetCellValue(sourceCell.NumericCellValue);
                                break;
                            case CellType.String:
                                targetCell.SetCellValue(sourceCell.RichStringCellValue);
                                break;
                        }
                    }
                }
            }

        }

        /// <summary>
        /// 在工作表中添加分页符。
        /// </summary>
        /// <param name="sheet">要添加分页符的工作表。</param>
        /// <param name="dataMaxCount">数据的最大行数。</param>
        /// <param name="rowsPerPage">每页的数据行数。</param>
        /// <param name="pageSize">页大小，默认为10。</param>
        public static void AddPageBreaks(this ISheet sheet, int dataMaxCount, int rowsPerPage, int pageSize = 10)
        {
            if (dataMaxCount <= pageSize)
            {
                var numberOfPages = dataMaxCount % pageSize == 0 ? dataMaxCount / pageSize : dataMaxCount / pageSize + 1;

                // 循环复制内容并添加分页符
                for (int i = 1; i < numberOfPages; i++)
                {
                    // 计算分页符的位置，即每一页结束的地方
                    int pageBreakRow = i * rowsPerPage - 1;
                    // 在该位置添加分页符
                    sheet.SetRowBreak(pageBreakRow);

                    // 计算新一页内容开始的行号
                    int startRow = pageBreakRow + 1;

                    // 复制第一页的内容到新的一页
                    for (int j = 0; j < rowsPerPage; j++)
                    {
                        // 获取源行，即第一页的行
                        IRow sourceRow = sheet.GetRow(j);
                        // 检查是否超出了工作表的现有行数，如果是，则创建新行
                        IRow targetRow = sheet.GetRow(startRow + j) ?? sheet.CreateRow(startRow + j);
                        // 调用复制行的方法来复制源行到目标行
                        CopyRow(sourceRow, targetRow);
                    }
                }
            }
        }



        /// <summary>
        /// 在工作表中设置单元格文字的值。
        /// </summary>
        /// <param name="sheet">要设置单元格值的工作表。</param>
        /// <param name="rowIndex">行索引。</param>
        /// <param name="cellIndex">列索引。</param>
        /// <param name="value">要设置的值。</param>
        public static void ExcelSetCellTextValue(this ISheet sheet, int rowIndex, int cellIndex, dynamic value)
        {
            IRow row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                row = sheet.CreateRow(rowIndex); // 如果行不存在，则创建新行
            }

            ICell cell = row.GetCell(cellIndex);
            if (cell == null)
            {
                cell = row.CreateCell(cellIndex);
            }

            cell.SetCellValue(Convert.ToString(value));
        }


        /// <summary>
        /// 在工作表中设置单元格的图片值。
        /// </summary>
        /// <param name="sheet">要设置图片值的工作表。</param>
        /// <param name="workBook">工作簿对象。</param>
        /// <param name="rowIndex">行索引。</param>
        /// <param name="cellIndex">列索引。</param>
        /// <param name="value">图片路径。</param>
        /// <param name="scaleX">水平缩放比例，默认为0。</param>
        /// <param name="scaleY">垂直缩放比例，默认为0。</param>
        public static void ExcelSetCellImageValue(this ISheet sheet, IWorkbook workBook, int rowIndex, int cellIndex, dynamic value, double scaleX = 0, double scaleY = 0)
        {
            // 加载图片
            byte[] imageBytes = File.ReadAllBytes(value);
            int pictureIndex = workBook.AddPicture(imageBytes, PictureType.JPEG);

            // 创建绘图对象
            IDrawing drawing;
            if (workBook is XSSFWorkbook xssfWorkbook)
            {
                drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;
            }
            else if (workBook is HSSFWorkbook hssfWorkbook)
            {
                drawing = sheet.CreateDrawingPatriarch() as HSSFPatriarch;
            }
            else
            {
                throw new ArgumentException("Unsupported workbook type.");
            }

            // 创建锚点
            IClientAnchor anchor;
            if (workBook is XSSFWorkbook)
            {
                anchor = new XSSFClientAnchor();
            }
            else
            {
                anchor = new HSSFClientAnchor();
            }
            anchor.Col1 = cellIndex;
            anchor.Row1 = rowIndex;
            anchor.Col2 = cellIndex + 1;
            anchor.Row2 = rowIndex + 1;

            // 在指定位置添加图片
            IPicture picture = drawing.CreatePicture(anchor, pictureIndex);

            // 调整图片大小（可选）
            if (scaleX != 0 && scaleY != 0)
            {
                picture.Resize(scaleX, scaleY);
            }
        }

    }
}
