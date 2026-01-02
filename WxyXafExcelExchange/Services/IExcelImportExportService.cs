using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using DevExpress.ExpressApp;

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// Excel导入导出核心服务接口（平台无关）
    /// </summary>
    public interface IExcelImportExportService
    {
        /// <summary>
        /// 导出数据到Excel/CSV
        /// </summary>
        /// <param name="data">要导出的数据</param>
        /// <param name="objectType">对象类型</param>
        /// <param name="options">导出选项</param>
        /// <returns>导出结果</returns>
        Task<ExcelExportResult> ExportDataAsync(IEnumerable<object> data, Type objectType, ExcelExportOptions options = null);

        /// <summary>
        /// 从Excel/CSV导入数据
        /// </summary>
        /// <param name="fileContent">文件内容</param>
        /// <param name="objectType">目标对象类型</param>
        /// <param name="objectSpace">XAF对象空间</param>
        /// <param name="options">导入选项</param>
        /// <param name="fileName">文件名(可选,用于格式检测)</param>
        /// <returns>导入结果</returns>
        Task<ExcelImportResult> ImportDataAsync(byte[] fileContent, Type objectType, IObjectSpace objectSpace, ExcelImportOptions options = null, string fileName = null);

        /// <summary>
        /// 验证文件格式
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="fileContent">文件内容</param>
        /// <returns>验证结果</returns>
        FileValidationResult ValidateFile(string fileName, byte[] fileContent);

        /// <summary>
        /// 获取对象类型的可导入/导出属性
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <returns>属性列表</returns>
        List<ExcelFieldInfo> GetExcelFields(Type objectType);

        /// <summary>
        /// 导出多个数据集到单个 Excel 文件的多个 Sheet (仅支持 XLSX 格式)
        /// </summary>
        /// <typeparam name="T">数据类型</typeparam>
        /// <param name="filePath">文件路径</param>
        /// <param name="sheetsData">Sheet 名称到数据的映射</param>
        /// <param name="headers">表头列表</param>
        /// <param name="dataExtractor">数据提取器函数</param>
        /// <param name="title">导出标题</param>
        /// <returns>导出任务</returns>
        Task ExportMultipleSheetsAsync<T>(
            string filePath,
            Dictionary<string, List<T>> sheetsData,
            string[] headers,
            Func<T, object[]> dataExtractor,
            string title = null);

        /// <summary>
        /// 获取 Excel 文件中所有 Sheet 的数据
        /// </summary>
        /// <param name="filePath">Excel 文件路径</param>
        /// <returns>Sheet 名称到数据列表的映射</returns>
        Task<Dictionary<string, List<Dictionary<string, object>>>> GetExcelSheetsAsync(string filePath);
    }

    /// <summary>
    /// Excel导出选项
    /// </summary>
    public class ExcelExportOptions
    {
        /// <summary>
        /// 导出格式（默认为 XLSX 以支持多 Sheet 功能）
        /// </summary>
        public ExcelFormat Format { get; set; } = ExcelFormat.Xlsx;

        /// <summary>
        /// 是否包含表头
        /// </summary>
        public bool IncludeHeaders { get; set; } = true;

        /// <summary>
        /// 文件名（不含扩展名）
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 编码格式（统一使用UTF-8）
        /// </summary>
        public System.Text.Encoding Encoding { get; set; } = System.Text.Encoding.UTF8;

        /// <summary>
        /// 是否添加BOM（UTF-8 BOM确保兼容性）
        /// </summary>
        public bool AddBom { get; set; } = true;

        /// <summary>
        /// 主表 Sheet 名称（仅 XLSX 格式有效）
        /// </summary>
        public string SheetName { get; set; }
    }

    /// <summary>
    /// Excel导入选项
    /// </summary>
    public class ExcelImportOptions
    {
        /// <summary>
        /// 导入模式
        /// </summary>
        public ExcelImportMode Mode { get; set; } = ExcelImportMode.CreateOrUpdate;

        /// <summary>
        /// 是否用户显式指定了导入模式(如果为true,则不应用类配置的默认模式)
        /// </summary>
        public bool IsUserSpecifiedMode { get; set; } = false;

        /// <summary>
        /// 是否跳过重复数据
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

        /// <summary>
        /// 批处理大小
        /// </summary>
        public int BatchSize { get; set; } = 1000;
    }

    /// <summary>
    /// Excel导出结果
    /// </summary>
    public class ExcelExportResult
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 导出的记录数
        /// </summary>
        public int RecordCount { get; set; }

        /// <summary>
        /// 文件内容
        /// </summary>
        public byte[] FileContent { get; set; }

        /// <summary>
        /// 建议的文件名
        /// </summary>
        public string SuggestedFileName { get; set; }

        /// <summary>
        /// MIME类型
        /// </summary>
        public string MimeType { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// Excel导入结果
    /// </summary>
    public class ExcelImportResult
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 总记录数
        /// </summary>
        public int TotalRecords { get; set; }

        /// <summary>
        /// 成功导入的记录数
        /// </summary>
        public int SuccessCount { get; set; }

        /// <summary>
        /// 失败的记录数
        /// </summary>
        public int FailureCount { get; set; }

        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 错误列表
        /// </summary>
        public List<ExcelImportError> Errors { get; set; } = new List<ExcelImportError>();

        /// <summary>
        /// 警告列表
        /// </summary>
        public List<ExcelImportWarning> Warnings { get; set; } = new List<ExcelImportWarning>();
    }

    /// <summary>
    /// 文件验证结果
    /// </summary>
    public class FileValidationResult
    {
        /// <summary>
        /// 是否有效
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// 检测到的格式
        /// </summary>
        public ExcelFormat DetectedFormat { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 文件大小（字节）
        /// </summary>
        public long FileSize { get; set; }
    }

    /// <summary>
    /// Excel字段信息
    /// </summary>
    public class ExcelFieldInfo
    {
        /// <summary>
        /// 属性名
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// 显示名称
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// 数据类型
        /// </summary>
        public Type DataType { get; set; }

        /// <summary>
        /// 是否必填
        /// </summary>
        public bool IsRequired { get; set; }

        /// <summary>
        /// 是否可导入
        /// </summary>
        public bool CanImport { get; set; }

        /// <summary>
        /// 是否可导出
        /// </summary>
        public bool CanExport { get; set; }

        /// <summary>
        /// 排序顺序
        /// </summary>
        public int Order { get; set; }
    }

    /// <summary>
    /// Excel导入错误
    /// </summary>
    public class ExcelImportError
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

        /// <summary>
        /// 错误类型
        /// </summary>
        public ExcelImportErrorType ErrorType { get; set; }
    }

    /// <summary>
    /// Excel导入警告
    /// </summary>
    public class ExcelImportWarning
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
        /// 警告消息
        /// </summary>
        public string WarningMessage { get; set; }

        /// <summary>
        /// 原始值
        /// </summary>
        public string OriginalValue { get; set; }

        /// <summary>
        /// 转换后的值
        /// </summary>
        public string ConvertedValue { get; set; }
    }

    /// <summary>
    /// Excel格式
    /// </summary>
    public enum ExcelFormat
    {
        /// <summary>
        /// CSV格式
        /// </summary>
        Csv,

        /// <summary>
        /// Excel 2007+ (.xlsx)
        /// </summary>
        Xlsx,

        /// <summary>
        /// Excel 97-2003 (.xls)
        /// </summary>
        Xls
    }

    /// <summary>
    /// Excel导入模式
    /// </summary>
    public enum ExcelImportMode
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
        ReplaceAll
    }

    /// <summary>
    /// Excel导入错误类型
    /// </summary>
    public enum ExcelImportErrorType
    {
        /// <summary>
        /// 数据类型转换错误
        /// </summary>
        DataTypeConversion,

        /// <summary>
        /// 必填字段为空
        /// </summary>
        RequiredFieldEmpty,

        /// <summary>
        /// 字段不存在
        /// </summary>
        FieldNotFound,

        /// <summary>
        /// 数据验证失败
        /// </summary>
        ValidationFailed,

        /// <summary>
        /// 重复数据
        /// </summary>
        DuplicateData,

        /// <summary>
        /// 系统错误
        /// </summary>
        SystemError
    }
}