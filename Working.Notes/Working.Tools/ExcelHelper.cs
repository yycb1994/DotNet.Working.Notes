using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;

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

        /// <summary>
        /// TODO:先创建行，然后在创建对应的列
        /// 创建Excel中指定的行
        /// </summary>
        /// <param name="sheet">Excel工作表对象</param>
        /// <param name="rowNum">创建第几行(从0开始)</param>
        /// <param name="rowHeight">行高</param>
        public static IRow CreateRow(ISheet sheet, int rowNum, float rowHeight)
        {
            IRow row = sheet.CreateRow(rowNum); //创建行
            row.HeightInPoints = rowHeight; //设置列头行高
            return row;
        }

        /// <summary>
        /// 创建行内指定的单元格
        /// </summary>
        /// <param name="row">需要创建单元格的行</param>
        /// <param name="cellStyle">单元格样式</param>
        /// <param name="cellNum">创建第几个单元格(从0开始)</param>
        /// <param name="cellValue">给单元格赋值</param>
        /// <returns></returns>
        public static ICell CreateCells(IRow row, ICellStyle cellStyle, int cellNum, string cellValue)
        {

            var cell = row.CreateCell(cellNum); //创建单元格
            if (cellStyle != null)
            {
                cell.CellStyle = cellStyle; //将样式绑定到单元格
            }

            if (!string.IsNullOrWhiteSpace(cellValue))
            {
                //单元格赋值
                cell.SetCellValue(cellValue);
            }

            return cell;
        }


        /// <summary>
        /// 行内单元格常用样式设置
        /// </summary>
        /// <param name="workbook">Excel文件对象</param>
        /// <param name="hAlignment">水平布局方式</param>
        /// <param name="vAlignment">垂直布局方式</param>
        /// <param name="fontHeightInPoints">字体大小</param>
        /// <param name="isAddBorder">是否需要边框</param>
        /// <param name="boldWeight">字体加粗 (None = 0,Normal = 400，Bold = 700</param>
        /// <param name="fontName">字体（仿宋，楷体，宋体，微软雅黑...与Excel主题字体相对应）</param>
        /// <param name="isAddBorderColor">是否增加边框颜色</param>
        /// <param name="isItalic">是否将文字变为斜体</param>
        /// <param name="isLineFeed">是否自动换行</param>
        /// <param name="isAddCellBackground">是否增加单元格背景颜色</param>
        /// <param name="fillPattern">填充图案样式(FineDots 细点，SolidForeground立体前景，isAddFillPattern=true时存在)</param>
        /// <param name="cellBackgroundColor">单元格背景颜色（当isAddCellBackground=true时存在）</param>
        /// <param name="fontColor">字体颜色</param>
        /// <param name="underlineStyle">下划线样式（无下划线[None],单下划线[Single],双下划线[Double],会计用单下划线[SingleAccounting],会计用双下划线[DoubleAccounting]）</param>
        /// <param name="typeOffset">字体上标下标(普通默认值[None],上标[Sub],下标[Super]),即字体在单元格内的上下偏移量</param>
        /// <param name="isStrikeout">是否显示删除线</param>
        /// <returns></returns>
        public static ICellStyle CreateStyle(IWorkbook workbook, HorizontalAlignment hAlignment, VerticalAlignment vAlignment, short fontHeightInPoints, bool isAddBorder, short boldWeight, string fontName = "宋体", bool isAddBorderColor = true, bool isItalic = false, bool isLineFeed = false, bool isAddCellBackground = false, FillPattern fillPattern = FillPattern.NoFill, short cellBackgroundColor = HSSFColor.Yellow.Index, short fontColor = HSSFColor.Black.Index, FontUnderlineType underlineStyle =
            FontUnderlineType.None, FontSuperScript typeOffset = FontSuperScript.None, bool isStrikeout = false)
        {
            var cellStyle = workbook.CreateCellStyle(); //创建列头单元格实例样式
            cellStyle.Alignment = hAlignment; //水平居中
            cellStyle.VerticalAlignment = vAlignment; //垂直居中
            cellStyle.WrapText = isLineFeed;//自动换行

            if (isAddCellBackground)
            {
                cellStyle.FillForegroundColor = cellBackgroundColor;//单元格背景颜色
                cellStyle.FillPattern = fillPattern;//填充图案样式(FineDots 细点，SolidForeground立体前景)
            }


            //是否增加边框
            if (isAddBorder)
            {
                //常用的边框样式 None(没有),Thin(细边框，瘦的),Medium(中等),Dashed(虚线),Dotted(星罗棋布的),Thick(厚的),Double(双倍),Hair(头发)[上右下左顺序设置]
                cellStyle.BorderBottom = BorderStyle.Thin;
                cellStyle.BorderRight = BorderStyle.Thin;
                cellStyle.BorderTop = BorderStyle.Thin;
                cellStyle.BorderLeft = BorderStyle.Thin;
            }

            //是否设置边框颜色
            if (isAddBorderColor)
            {
                //边框颜色[上右下左顺序设置]
                cellStyle.TopBorderColor = HSSFColor.DarkGreen.Index;//DarkGreen(黑绿色)
                cellStyle.RightBorderColor = HSSFColor.DarkGreen.Index;
                cellStyle.BottomBorderColor = HSSFColor.DarkGreen.Index;
                cellStyle.LeftBorderColor = HSSFColor.DarkGreen.Index;
            }

            /**
             * 设置相关字体样式
             */
            var cellStyleFont = workbook.CreateFont(); //创建字体

            //假如字体大小只需要是粗体的话直接使用下面该属性即可
            //cellStyleFont.IsBold = true;

            cellStyleFont.Boldweight = boldWeight; //字体加粗
            cellStyleFont.FontHeightInPoints = fontHeightInPoints; //字体大小
            cellStyleFont.FontName = fontName;//字体（仿宋，楷体，宋体 ）
            cellStyleFont.Color = fontColor;//设置字体颜色
            cellStyleFont.IsItalic = isItalic;//是否将文字变为斜体
            cellStyleFont.Underline = underlineStyle;//字体下划线
            cellStyleFont.TypeOffset = typeOffset;//字体上标下标
            cellStyleFont.IsStrikeout = isStrikeout;//是否有删除线

            cellStyle.SetFont(cellStyleFont); //将字体绑定到样式
            return cellStyle;
        }

        /// <summary>
        /// 一个用来添加分页符的函数
        /// 每隔pageSize行去插入一个分页符，第二页以后的所有内容，都去复制第一页的 numberOfPages 决定要插入多少页
        /// </summary>
        /// <param name="inputFilePath"></param>
        /// <param name="outputFilePath"></param>
        /// <param name="pageSize"></param>
        /// <param name="numberOfPages"></param>
        /// <param name="dataSource">数据源</param>
        public static void AddPageBreaks(string inputFilePath, string outputFilePath, int pageSize, int numberOfPages)
        {
            IWorkbook workbook;
            // 打开一个文件流来读取Excel文件
            using (FileStream file = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read))
            {
                // 根据文件扩展名决定使用哪个类来处理Excel文件
                if (Path.GetExtension(inputFilePath).ToLower() == ".xlsx")
                    workbook = new XSSFWorkbook(file);
                else
                    workbook = new HSSFWorkbook(file);
            }

            // 获取Excel文件的第一个工作表
            ISheet sheet = workbook.GetSheetAt(0);

            // 设置每页的行数
            int rowsPerPage = pageSize;

            // 循环添加分页符并复制内容
            for (int i = 1; i < numberOfPages; i++)
            {
                // 计算分页符的位置
                int pageBreakRow = i * rowsPerPage - 1;
                // 在计算出的行上方添加分页符
                sheet.SetRowBreak(pageBreakRow);

                // 复制第一页的内容到新的一页
                for (int j = 0; j < rowsPerPage; j++)
                {
                    // 获取源行
                    IRow sourceRow = sheet.GetRow(j);
                    // 创建新的目标行
                    IRow targetRow = sheet.CreateRow(pageBreakRow + 1 + j);
                    // 复制源行到目标行
                    CopyRow(workbook, sourceRow, targetRow);
                }
            }

            // 创建一个文件流来写入修改后的Excel文件
            using (FileStream outFile = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write))
            {
                // 将工作簿的内容写入文件流，即保存文件
                workbook.Write(outFile);
            }

            // 打开保存后的文件，前提是系统关联了Excel程序来打开.xls文件
           // System.Diagnostics.Process.Start(outputFilePath);
        }

        // 用于复制行的函数
        private static void CopyRow(IWorkbook workbook, IRow sourceRow, IRow targetRow)
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


        public static void AddPageBreaks1(string inputFilePath, string outputFilePath, int pageSize, int numberOfPages)
        {
            IWorkbook workbook;
            // 打开一个文件流来读取Excel文件
            using (FileStream file = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read))
            {
                // 根据文件扩展名决定使用哪个类来处理Excel文件
                if (Path.GetExtension(inputFilePath).ToLower() == ".xlsx")
                    workbook = new XSSFWorkbook(file);
                else
                    workbook = new HSSFWorkbook(file);
            }

            // 获取Excel文件的第一个工作表
            ISheet sheet = workbook.GetSheetAt(0);

            // 设置每页的行数
            int rowsPerPage = pageSize;

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
                    CopyRow(workbook, sourceRow, targetRow);
                }
            }

          

            // 打开保存后的文件，前提是系统关联了Excel程序来打开.xls文件
            // System.Diagnostics.Process.Start(outputFilePath);
        }


    }
}
