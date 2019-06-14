using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using 降水侵蚀力计算.Model;

namespace 降水侵蚀力计算.Operator
{
    /// <summary>
    /// Excel 操作基类
    /// </summary>
    public abstract class BaseOperator
    {
        /// <summary>
        /// 获取 Excel 文件数据表格
        /// </summary>
        internal ExcelDataGrid ExcelDataGrid { get; private set; }

        /// <summary>
        /// 打开指定的 Excel 文件
        /// </summary>
        /// <param name="fileName">指定要打开的 Excel 文件名</param>
        /// <param name="sheetIndex">指定要打开的 Sheet 索引</param>
        /// <returns>Excel 文件数据表格</returns>
        internal ExcelDataGrid Open(string fileName, int sheetIndex = 0)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentNullException(nameof(fileName));
            }
            var sourceFs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            var extensions = new FileInfo(fileName).Extension;
            IWorkbook workbook = null;
            IFormulaEvaluator evaluator = null;
            if (extensions.Equals(".xlsx", StringComparison.CurrentCultureIgnoreCase))
            {
                workbook = new XSSFWorkbook(sourceFs);
                evaluator = new XSSFFormulaEvaluator(workbook);
            }
            else if (extensions.Equals(".xls", StringComparison.CurrentCultureIgnoreCase))
            {
                workbook = new HSSFWorkbook(sourceFs);
                evaluator = new HSSFFormulaEvaluator(workbook);
            }
            else
            {
                throw new ArgumentException("不支持的文件扩展名");
            }
            var sheet = workbook.GetSheetAt(sheetIndex);
            if (sheet == null)
                throw new FileLoadException("打开的 Excel 文件没有 Sheet.");

            var result = new ExcelDataGrid();
            var rowIndex = -1;
            var columnIndex = -1;
            foreach (var row in GetRows(sheet))
            {
                rowIndex++;
                columnIndex = -1;
                foreach (var column in row.Cells)
                {
                    columnIndex++;
                    result[rowIndex, columnIndex] = GetCellValue(evaluator, column);
                }
            }
            sourceFs.Close();
            workbook.Close();
            return result;
        }

        /// <summary>
        /// 将数据写入到指定的 Excel 文件中
        /// </summary>
        /// <param name="targetFileName">指定的文件名</param>
        /// <param name="sheetIndex">指定要打开的 Sheet 索引</param>
        /// <param name="mergeSamColumn">同一行中相邻列的值相同, 则合并</param>
        /// <param name="mergeSameRow">同一列中相邻行的值相同, 则合并</param>
        /// <param name="dataGrid">指定的数据</param>
        internal void Write(string targetFileName, string templateFileName, ExcelDataGrid dataGrid, int sheetIndex = 0, string sheetName = null, bool mergeSameRow = false, bool mergeSamColumn = false)
        {
            if (string.IsNullOrWhiteSpace(targetFileName))
            {
                throw new ArgumentNullException(nameof(targetFileName));
            }
            var sourceFs = new FileStream(templateFileName, FileMode.Open, FileAccess.Read);
            var targetFs = new FileStream(targetFileName, FileMode.Create, FileAccess.ReadWrite);
            var extensions = new FileInfo(templateFileName).Extension;
            IWorkbook workbook;
            if (extensions.Equals(".xlsx", StringComparison.CurrentCultureIgnoreCase))
            {
                workbook = new XSSFWorkbook(sourceFs);
            }
            else if (extensions.Equals(".xls", StringComparison.CurrentCultureIgnoreCase))
            {
                workbook = new HSSFWorkbook(sourceFs);
            }
            else
            {
                throw new ArgumentException("不支持的文件扩展名");
            }
            ISheet sheet;
            try
            {
                sheet = workbook.GetSheetAt(sheetIndex);
                if (!string.IsNullOrEmpty(sheetName))
                    workbook.SetSheetName(sheetIndex, sheetName);
            }
            catch (ArgumentException)
            {
                if (string.IsNullOrEmpty(sheetName))
                    sheet = workbook.CreateSheet();
                else
                    sheet = workbook.CreateSheet(sheetName);
            }
            var rowIndex = -1;
            foreach (var row in dataGrid.Rows)
            {
                rowIndex++;
                var columnIndex = -1;
                foreach (var column in row.Columns)
                {
                    columnIndex++;
                    var value = column;
                    if (value != null && value.Value != null && value != Model.CellValue.Skip)
                    {
                        var excelRow = sheet.GetRow(rowIndex);
                        if (excelRow == null)
                            excelRow = sheet.CreateRow(rowIndex);
                        var excelColumn = excelRow.GetCell(columnIndex);
                        if (excelColumn == null)
                            excelColumn = excelRow.CreateCell(columnIndex);
                        if (excelColumn.CellStyle == null)
                            excelColumn.CellStyle = workbook.CreateCellStyle();
                        excelColumn.CellStyle.Alignment = HorizontalAlignment.Center;
                        excelColumn.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        var @string = value.Value as string;
                        if (@string != null)
                        {
                            excelColumn.SetCellValue(@string);
                        }
                        var @bool = value.Value as bool?;
                        if (@bool != null)
                        {
                            excelColumn.SetCellValue(@bool.Value);
                        }
                        var dateTime = value.Value as DateTime?;
                        if (dateTime != null)
                        {
                            excelColumn.SetCellValue(dateTime.Value);
                            var cellStyle = workbook.CreateCellStyle();
                            var format = workbook.CreateDataFormat();
                            cellStyle.Alignment = HorizontalAlignment.Center;
                            cellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cellStyle.DataFormat = format.GetFormat(value.Format ?? "yyyy-MM-dd HH:mm:ss");
                            excelColumn.CellStyle = cellStyle;
                            sheet.AutoSizeColumn(columnIndex);
                        }
                        if (value.Value is double
                           || value.Value is float
                           || value.Value is int
                           || value.Value is decimal)
                        {
                            excelColumn.SetCellValue((double)value.Value);
                        }
                    }

                }
            }

            if (mergeSameRow)
            {
                var columnCount = dataGrid.Rows.Max(p => p.Columns.Count);
                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    var firstRowIndex = 0;
                    var lastRowIndex = 0;
                    for (rowIndex = 0; rowIndex < dataGrid.Rows.Count; rowIndex++)
                    {
                        var curtRow = dataGrid.Rows[rowIndex].Columns[columnIndex];
                        if (curtRow == null || curtRow.Value == null || string.IsNullOrEmpty(curtRow.Value as string))
                        {
                            lastRowIndex = rowIndex;
                        }
                        else
                        {
                            if (lastRowIndex - firstRowIndex > 0)
                            {
                                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(firstRowIndex, lastRowIndex, columnIndex, columnIndex));
                            }
                            firstRowIndex = rowIndex;
                        }

                    }
                }
            }

            workbook.Write(targetFs);
            sourceFs.Close();
            targetFs.Close();
            workbook.Close();
        }


        #region - Private -
        /// <summary>
        /// 获取<see cref="ISheet"/>中的行
        /// </summary>
        /// <param name="sheet">指定获取的<see cref="ISheet"/>实例</param>
        /// <returns><see cref="ISheet"/>中的行</returns>
        private IEnumerable<IRow> GetRows(ISheet sheet)
        {
            var enumerator = sheet.GetEnumerator();
            while (enumerator.MoveNext())
            {
                yield return enumerator.Current as IRow;
            }
        }

        /// <summary>
        /// 获取单元格的显示值
        /// </summary>
        /// <param name="evaluator">单元格公式计算器</param>
        /// <param name="cell">单元格</param>
        /// <returns>单元格显示的值</returns>
        private string GetCellValue(IFormulaEvaluator evaluator, ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Blank:
                    return string.Empty;
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
                case CellType.Formula:
                    cell = evaluator.EvaluateInCell(cell);
                    return GetCellValue(evaluator, cell);
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                        return cell.DateCellValue.ToString();
                    else
                        return cell.NumericCellValue.ToString();
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Unknown:
                    return "Unknow";
                default:
                    return "";
            }
        }
        #endregion
    }
}
