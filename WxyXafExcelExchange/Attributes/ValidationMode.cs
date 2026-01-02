using System;

namespace Wxy.Xaf.ExcelExchange.Enums
{
    /// <summary>
    /// 验证模式
    /// </summary>
    public enum ValidationMode
    {
        /// <summary>
        /// 严格模式：任意行错误则终止导入
        /// </summary>
        Strict,

        /// <summary>
        /// 宽松模式：跳过错误行，记录错误信息
        /// </summary>
        Loose
    }
}