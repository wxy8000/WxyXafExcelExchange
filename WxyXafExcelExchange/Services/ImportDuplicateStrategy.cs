using System;

namespace Wxy.Xaf.ExcelExchange.Enums
{
    /// <summary>
    /// 重复数据处理策略
    /// </summary>
    public enum ImportDuplicateStrategy
    {
        /// <summary>
        /// 仅新增，重复则报错
        /// </summary>
        Insert,

        /// <summary>
        /// 仅更新，无则报错
        /// </summary>
        Update,

        /// <summary>
        /// 新增或更新（根据匹配字段）
        /// </summary>
        InsertOrUpdate,

        /// <summary>
        /// 忽略重复行
        /// </summary>
        Ignore
    }
}