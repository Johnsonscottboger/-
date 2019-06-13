using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using 降水侵蚀力计算.Model;
using 降水侵蚀力计算.Operator;

namespace 降水侵蚀力计算.Handler
{
    /// <summary>
    /// 降水侵蚀力计算
    /// </summary>
    internal class DefaultHandler : BaseOperator, IHandler
    {
        /// <summary>
        /// 获取文件名
        /// </summary>
        public string FileName { get; private set; }

        /// <summary>
        /// 初始化<see cref="DefaultHandler"/>实例
        /// </summary>
        /// <param name="fileName">指定的文件名</param>
        public DefaultHandler(string fileName)
        {
            this.FileName = fileName ?? throw new ArgumentNullException(nameof(fileName));
        }

        /// <summary>
        /// 处理
        /// </summary>
        public void Handle()
        {
            var data = Open(this.FileName, 1);
            var records = Load(data);

            RinfallPartition(records);

            RainfallDaySum(records);
        }

        #region - Private -

        /// <summary>
        /// A.划分次降水
        /// </summary>
        /// <param name="records">降水记录</param>
        private List<PartitionRecord> RinfallPartition(IEnumerable<Record> records)
        {
            //1. 降水间歇 > 6h, 分为两次降水
            //2. 降水没有间歇, 但是6之后降水量< 1.3 mm, 分为两次降水
            var list = new List<PartitionRecord>();

            //降水序号
            var number = 0;
            //是否正在降水
            var raining = false;

            //首次降水时间
            var firstRain = default(DateTime?);
            //首次非降水时间
            var firstURain = default(DateTime?);

            //最近降水时间
            var lastRain = default(DateTime?);
            
            foreach (var record in records)
            {
                //降雨
                if (record.Precipitation30 > 0)
                {
                    lastRain = record.DateTime;
                    if (!raining)
                    {
                        firstRain = record.DateTime;
                        if ((firstRain - firstURain).Value.TotalHours > 6)
                            number = number == 0 ? 1 : number + 1;
                        raining = true;
                    }
                    else
                    {
                        var h = (lastRain - firstRain).Value.TotalHours;
                        if (h >= 6)
                        {
                            var sum = list.Where(p => p.Number == number).Sum(p => p.Precipitation30);
                            if (sum > 1.3m)
                            {
                                number++;
                            }
                        }
                    }
                    list.Add(new PartitionRecord()
                    {
                        Number = number,
                        DateTime = record.DateTime,
                        Precipitation15 = record.Precipitation15,
                        Precipitation30 = record.Precipitation30
                    });
                }
                //未降雨
                else
                {
                    if (raining)
                    {
                        raining = false;
                        firstURain = record.DateTime;
                    }
                    else if (firstURain == null)
                    {
                        firstURain = record.DateTime;
                    }
                }
            }

            var dataGrid = new ExcelDataGrid();
            dataGrid[0, 0] = "序号";
            dataGrid[0, 1] = "时间";
            dataGrid[0, 2] = "I30";
            dataGrid[0, 3] = "I15";
            dataGrid[0, 4] = "降水量";
            var rowIndex = 0;
            var lastNumber = 0;
            foreach (var item in list)
            {
                rowIndex++;
                dataGrid[rowIndex, 0] = item.Number == lastNumber ? CellValue.Skip : item.Number.ToString();
                dataGrid[rowIndex, 1] = item.DateTime.ToString("yyyy-MM-dd HH:mm");
                dataGrid[rowIndex, 2] = item.Precipitation30.ToString();
                dataGrid[rowIndex, 3] = item.Precipitation15.ToString();
                dataGrid[rowIndex, 4] = item.Number == lastNumber ? CellValue.Skip : list.Where(p => p.Number == item.Number).Sum(p => p.Precipitation30).ToString();
                lastNumber = item.Number;
            }
            Fill(dataGrid, 2);
            return list;
        }

        /// <summary>
        /// B.侵蚀性降水筛选
        /// </summary>
        /// <param name="records">划分次降水记录</param>
        private void RainfallFilte(IEnumerable<PartitionRecord> records)
        {

        }

        /// <summary>
        /// C.统计日降水量
        /// </summary>
        /// <param name="records">降水记录</param>
        private void RainfallDaySum(IEnumerable<Record> records)
        {
            var list = new List<DaySumRecord>();

            records.Where(p => p.Precipitation30 > 0)
                   .GroupBy(p => p.DateTime.Date)
                   .Select(p =>
                   {
                       var sum = p.Sum(c => c.Precipitation30);
                       foreach(var item in p)
                       {
                           list.Add(new DaySumRecord
                           {
                               Date = p.Key,
                               DateTime = item.DateTime,
                               Precipitation15 = item.Precipitation15,
                               Precipitation30 = item.Precipitation30,
                               PrecipitationSum = sum
                           });
                       }
                       return 0;
                   }).ToList();

            var dataGrid = new ExcelDataGrid();
            dataGrid[0, 0] = "日期";
            dataGrid[0, 1] = "时间";
            dataGrid[0, 2] = "I30";
            dataGrid[0, 3] = "I15";
            dataGrid[0, 4] = "降水量";
            var rowIndex = 0;
            var lastDate = new DateTime();
            foreach(var item in list)
            {
                rowIndex++;
                dataGrid[rowIndex, 0] = item.Date == lastDate ? CellValue.Skip : item.Date.ToString("yyyy-MM-dd");
                dataGrid[rowIndex, 1] = item.DateTime.ToString("yyyy-MM-dd HH:mm");
                dataGrid[rowIndex, 2] = item.Precipitation30.ToString();
                dataGrid[rowIndex, 3] = item.Precipitation15.ToString();
                dataGrid[rowIndex, 4] = item.Date == lastDate ? CellValue.Skip : item.PrecipitationSum.ToString();
                lastDate = item.Date;
            }

            Fill(dataGrid, 3);
        }
        #endregion

        #region - 基础方法 -

        /// <summary>
        /// 加载记录, 将<see cref="ExcelDataGrid"/>转换为<see cref="Record"/>
        /// </summary>
        private IEnumerable<Record> Load(ExcelDataGrid dataGrid)
        {
            var list = new List<Record>();
            if (dataGrid == null)
                return list;
            foreach (var row in dataGrid.Rows)
            {
                if (DateTime.TryParse(row.Columns[0].Value, out var dateTime)
                    && decimal.TryParse(row.Columns[1].Value, out var p30)
                    && decimal.TryParse(row.Columns[2].Value, out var p15))
                {
                    list.Add(new Record()
                    {
                        DateTime = dateTime,
                        Precipitation30 = p30,
                        Precipitation15 = p15
                    });
                }
            }
            return list;
        }


        private void Fill(ExcelDataGrid dataGrid, int sheetIndex)
        {
            var targetFileName = this.FileName;
            var fi = new FileInfo(targetFileName);
            var templateFileName = $"{fi.Name} - Template{fi.Extension}";
            File.Copy(targetFileName, templateFileName, true);
            Write(targetFileName, templateFileName, dataGrid, sheetIndex, true);
            File.Delete(templateFileName);
        }

        /// <summary>
        /// 记录
        /// </summary>
        private class Record
        {
            /// <summary>
            /// 获取或设置降水时间
            /// </summary>
            public DateTime DateTime { get; set; }

            /// <summary>
            /// 获取或设置15分钟降水量
            /// </summary>
            public decimal Precipitation15 { get; set; }

            /// <summary>
            /// 获取或设置30分钟降水量
            /// </summary>
            public decimal Precipitation30 { get; set; }

            public override string ToString()
            {
                return $"{{ DateTime:{this.DateTime.ToString()}, I15:{this.Precipitation15.ToString()}, I30:{this.Precipitation30.ToString()}}}";
            }
        }

        /// <summary>
        /// 划分次降雨
        /// </summary>
        private class PartitionRecord : Record
        {
            /// <summary>
            /// 获取或设置序号
            /// </summary>
            public int Number { get; set; }

            public override string ToString()
            {
                return $"{{Number: {this.Number.ToString()}, DateTime:{this.DateTime.ToString()}, I15:{this.Precipitation15.ToString()}, I30:{this.Precipitation30.ToString()}}}";
            }
        }

        /// <summary>
        /// 日降水量
        /// </summary>
        private class DaySumRecord : Record
        {
            /// <summary>
            /// 获取或设置日期
            /// </summary>
            public DateTime Date { get; set; }

            /// <summary>
            /// 获取或设置降水量合计
            /// </summary>
            public decimal PrecipitationSum { get; set; }
        }
        #endregion
    }
}
