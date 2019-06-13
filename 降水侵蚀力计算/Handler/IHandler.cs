using System;
using System.Collections.Generic;
using System.Text;

namespace 降水侵蚀力计算.Handler
{
    /// <summary>
    /// 处理器
    /// </summary>
    internal interface IHandler
    {
        /// <summary>
        /// 获取文件名
        /// </summary>
        string FileName { get;  }

        /// <summary>
        /// 执行
        /// </summary>
        void Handle();
    }
}
