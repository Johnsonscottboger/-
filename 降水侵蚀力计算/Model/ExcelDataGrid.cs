using System;
using System.Collections.Generic;
using System.Text;

namespace 降水侵蚀力计算.Model
{
    /// <summary>
    /// 表示数据表格
    /// </summary>
    internal class ExcelDataGrid
    {
        /// <summary>
        /// 表格中的行
        /// </summary>
        public List<Row> Rows { get; set; }

        /// <summary>
        /// 初始化<see cref="ExcelDataGrid"/>
        /// </summary>
        public ExcelDataGrid()
        {
            this.Rows = new List<Row>();
        }

        /// <summary>
        /// 根据指定索引获取或设置值
        /// </summary>
        /// <param name="rowIndex">指定的行索引</param>
        /// <param name="columnIndex">指定的列索引</param>
        /// <returns>获取到指定索引的值</returns>
        public CellValue this[int rowIndex, int columnIndex]
        {
            get
            {
                var row = this.Rows[rowIndex];
                var cell = row.Columns[columnIndex];
                return cell;
            }
            set
            {
                if (rowIndex >= this.Rows.Count)
                {
                    for (var i = this.Rows.Count; i <= rowIndex; i++)
                    {
                        this.Rows.Add(new Row());
                    }
                }
                var row = this.Rows[rowIndex];
                if (columnIndex >= row.Columns.Count)
                {
                    for (var i = row.Columns.Count; i <= columnIndex; i++)
                    {
                        row.Columns.Add(new CellValue());
                    }
                }
                row.Columns[columnIndex] = value;
            }
        }
    }

    /// <summary>
    /// 表示一行数据
    /// </summary>
    internal class Row
    {
        /// <summary>
        /// 获取或设置包括的列
        /// </summary>
        public List<CellValue> Columns { get; set; }

        /// <summary>
        /// 初始化<see cref="Row"/>实例
        /// </summary>
        public Row()
        {
            this.Columns = new List<CellValue>();
        }
    }

    /// <summary>
    /// 表示一个单元格数据
    /// </summary>
    internal class CellValue
    {
        /// <summary>
        /// 实际值
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// 获取或设置单元格格式
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 指示当前单元格为无效数据, 返回字符串 "-"
        /// </summary>
        public static CellValue Empty { get { return "-"; } }

        /// <summary>
        /// 指示当前单元格为空, 在填充时, 将跳过当前单元格的填充
        /// </summary>
        public static CellValue Skip { get; set; }

        /// <summary>
        /// 初始化默认单元格数据
        /// </summary>
        public CellValue() { }

        /// <summary>
        /// 使用指定实际值初始化单元格数据
        /// </summary>
        /// <param name="value">将初始化为单元格实际值</param>
        public CellValue(object value)
        {
            this.Value = value;
        }

        /// <summary>
        /// 初始化单元格数据
        /// </summary>
        /// <param name="value">指定的实际值</param>
        /// <param name="format">指定的显示格式</param>
        public CellValue(object value, string format)
        {
            this.Value = value;
            this.Format = format;
        }

        public static implicit operator CellValue(string value)
        {
            return new CellValue(value);
        }

        public static implicit operator CellValue(bool value)
        {
            return new CellValue(value);
        }

        public static implicit operator CellValue(DateTime value)
        {
            return new CellValue(value);
        }

        public static implicit operator CellValue(double value)
        {
            return new CellValue(value);
        }

        public static implicit operator CellValue(float value)
        {
            return new CellValue((double)value);
        }

        public static implicit operator CellValue(int value)
        {
            return new CellValue((double)value);
        }

        public static implicit operator CellValue(decimal value)
        {
            return new CellValue((double)value);
        }

        /// <summary>
        /// 返回表示当前对象的字符串
        /// </summary>
        /// <returns>表示当前对象的字符串</returns>
        public override string ToString()
        {
            return $"Value:{this.Value}";
        }
    }
}
