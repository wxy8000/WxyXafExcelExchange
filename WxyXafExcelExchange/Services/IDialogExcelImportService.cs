using System;
using System.Collections.Generic;
using System.IO;
using DevExpress.ExpressApp;

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// 对话框Excel导入服务接口
    /// </summary>
    public interface IDialogExcelImportService
    {
        /// <summary>
        /// 从URI解析对象类型
        /// </summary>
        /// <param name="uri">URI</param>
        /// <returns>对象类型</returns>
        Type ResolveObjectTypeFromUri(Uri uri);

        /// <summary>
        /// 解析导入模式
        /// </summary>
        /// <param name="modeString">模式字符串</param>
        /// <returns>导入模式</returns>
        XpoExcelImportMode ParseImportMode(string modeString);

        /// <summary>
        /// 从Excel流导入数据
        /// </summary>
        /// <param name="stream">Excel文件流</param>
        /// <param name="objectType">对象类型</param>
        /// <param name="options">导入选项</param>
        /// <param name="application">XAF应用程序实例</param>
        /// <returns>导入结果</returns>
        XpoExcelImportResult ImportFromExcelStream(Stream stream, Type objectType, XpoExcelImportOptions options, XafApplication application);
        
        /// <summary>
        /// 从CSV流导入数据
        /// </summary>
        /// <param name="stream">CSV文件流</param>
        /// <param name="objectType">对象类型</param>
        /// <param name="options">导入选项</param>
        /// <param name="application">XAF应用程序实例</param>
        /// <returns>导入结果</returns>
        XpoExcelImportResult ImportFromCsvStream(Stream stream, Type objectType, XpoExcelImportOptions options, XafApplication application);
    }

    /// <summary>
    /// Excel导入模式
    /// </summary>
    public enum XpoExcelImportMode
    {
        /// <summary>
        /// 仅创建新记录
        /// </summary>
        CreateOnly,

        /// <summary>
        /// 仅更新现有记录
        /// </summary>
        UpdateOnly,

        /// <summary>
        /// 创建或更新记录
        /// </summary>
        CreateOrUpdate,

        /// <summary>
        /// 替换所有记录
        /// </summary>
        Replace
    }

    /// <summary>
    /// Excel导入选项
    /// </summary>
    public class XpoExcelImportOptions
    {
        /// <summary>
        /// 导入模式
        /// </summary>
        public XpoExcelImportMode Mode { get; set; } = XpoExcelImportMode.CreateOrUpdate;

        /// <summary>
        /// 跳过重复数据
        /// </summary>
        public bool SkipDuplicates { get; set; } = true;

        /// <summary>
        /// 是否有表头行
        /// </summary>
        public bool HasHeaderRow { get; set; } = true;

        /// <summary>
        /// 最大允许错误数
        /// </summary>
        public int MaxErrors { get; set; } = 100;
    }

    /// <summary>
    /// Excel导入结果
    /// </summary>
    public class XpoExcelImportResult
    {
        /// <summary>
        /// 成功导入的记录数
        /// </summary>
        public int SuccessCount { get; set; }

        /// <summary>
        /// 失败的记录数
        /// </summary>
        public int FailedCount { get; set; }

        /// <summary>
        /// 错误列表
        /// </summary>
        public List<XpoExcelImportError> Errors { get; set; } = new List<XpoExcelImportError>();

        /// <summary>
        /// 是否成功
        /// </summary>
        public bool IsSuccess => FailedCount == 0;

        /// <summary>
        /// 总记录数
        /// </summary>
        public int TotalCount => SuccessCount + FailedCount;
    }

    /// <summary>
    /// Excel导入错误
    /// </summary>
    public class XpoExcelImportError
    {
        /// <summary>
        /// 行号
        /// </summary>
        public int RowNumber { get; set; }

        /// <summary>
        /// 字段名
        /// </summary>
        public string FieldName { get; set; }

        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 原始值
        /// </summary>
        public string OriginalValue { get; set; }
    }
}