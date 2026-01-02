using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// Excel导入服务接口
    /// </summary>
    public interface IImportService
    {
        /// <summary>
        /// 从CSV数据导入对象列表
        /// </summary>
        /// <typeparam name="T">目标对象类型</typeparam>
        /// <param name="csvData">CSV数据</param>
        /// <param name="options">导入选项</param>
        /// <returns>导入结果</returns>
        Task<ImportResult<T>> ImportFromCsvAsync<T>(string csvData, ImportOptions options = null) where T : class, new();

        /// <summary>
        /// 从CSV数据导入对象列表（非泛型版本）
        /// </summary>
        /// <param name="objectType">目标对象类型</param>
        /// <param name="csvData">CSV数据</param>
        /// <param name="options">导入选项</param>
        /// <returns>导入结果</returns>
        Task<ImportResult> ImportFromCsvAsync(Type objectType, string csvData, ImportOptions options = null);

        /// <summary>
        /// 验证CSV数据格式
        /// </summary>
        /// <param name="objectType">目标对象类型</param>
        /// <param name="csvData">CSV数据</param>
        /// <returns>验证结果</returns>
        HeaderValidationResult ValidateCsvFormat(Type objectType, string csvData);

        /// <summary>
        /// 获取对象类型的导入模板（CSV表头）
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <returns>CSV表头字符串</returns>
        string GetImportTemplate(Type objectType);

        /// <summary>
        /// 获取对象类型的可导入字段信息
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <returns>字段信息列表</returns>
        List<ImportFieldInfo> GetImportableFields(Type objectType);
    }

    /// <summary>
    /// 导入结果（泛型版本）
    /// </summary>
    /// <typeparam name="T">对象类型</typeparam>
    public class ImportResult<T> : ImportResult where T : class
    {
        /// <summary>
        /// 成功导入的对象列表
        /// </summary>
        public new List<T> SuccessfulObjects { get; set; } = new List<T>();
    }

    /// <summary>
    /// 导入结果
    /// </summary>
    public class ImportResult
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
        /// 成功记录数
        /// </summary>
        public int SuccessfulRecords { get; set; }

        /// <summary>
        /// 失败记录数
        /// </summary>
        public int FailedRecords { get; set; }

        /// <summary>
        /// 错误列表
        /// </summary>
        public List<ImportError> Errors { get; set; } = new List<ImportError>();

        /// <summary>
        /// 成功导入的对象列表（非泛型版本）
        /// </summary>
        public List<object> SuccessfulObjects { get; set; } = new List<object>();

        /// <summary>
        /// 处理耗时（毫秒）
        /// </summary>
        public long ProcessingTimeMs { get; set; }

        /// <summary>
        /// 添加错误
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnName">列名</param>
        /// <param name="message">错误消息</param>
        /// <param name="errorType">错误类型</param>
        public void AddError(int rowIndex, string columnName, string message, ImportErrorType errorType = ImportErrorType.ValidationError)
        {
            Errors.Add(new ImportError
            {
                RowIndex = rowIndex,
                ColumnName = columnName,
                Message = message,
                ErrorType = errorType
            });
        }

        /// <summary>
        /// 添加错误
        /// </summary>
        /// <param name="error">错误对象</param>
        public void AddError(ImportError error)
        {
            Errors.Add(error);
        }
    }

    /// <summary>
    /// 导入错误
    /// </summary>
    public class ImportError
    {
        /// <summary>
        /// 行索引（从1开始）
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 列名
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// 错误消息
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 错误类型
        /// </summary>
        public ImportErrorType ErrorType { get; set; }

        /// <summary>
        /// 原始值
        /// </summary>
        public string OriginalValue { get; set; }

        /// <summary>
        /// 异常信息
        /// </summary>
        public Exception Exception { get; set; }

        public override string ToString()
        {
            return $"第{RowIndex}行，列'{ColumnName}': {Message}";
        }
    }

    /// <summary>
    /// 导入模式枚举
    /// </summary>
    public enum ImportMode
    {
        /// <summary>
        /// 仅创建新记录，如果记录已存在则跳过
        /// </summary>
        CreateOnly,

        /// <summary>
        /// 仅更新现有记录，如果记录不存在则跳过
        /// </summary>
        UpdateOnly,

        /// <summary>
        /// 创建或更新：如果记录存在则更新，不存在则创建
        /// </summary>
        CreateOrUpdate,

        /// <summary>
        /// 替换所有数据：删除现有数据后导入新数据
        /// </summary>
        ReplaceAll
    }

    /// <summary>
    /// 导入错误类型
    /// </summary>
    public enum ImportErrorType
    {
        /// <summary>
        /// 解析错误
        /// </summary>
        ParseError,

        /// <summary>
        /// 验证错误
        /// </summary>
        ValidationError,

        /// <summary>
        /// 类型转换错误
        /// </summary>
        ConversionError,

        /// <summary>
        /// 关联对象查找错误
        /// </summary>
        LookupError,

        /// <summary>
        /// 系统错误
        /// </summary>
        SystemError
    }

    /// <summary>
    /// 导入选项
    /// </summary>
    public class ImportOptions
    {
        /// <summary>
        /// 导入模式，默认为CreateOrUpdate
        /// </summary>
        public ImportMode ImportMode { get; set; } = ImportMode.CreateOrUpdate;

        /// <summary>
        /// 重复检测的键属性名称列表（用于判断记录是否重复）
        /// </summary>
        public List<string> DuplicateDetectionKeys { get; set; } = new List<string>();

        /// <summary>
        /// 是否创建缺失的关联对象，默认为false
        /// </summary>
        public bool CreateMissingReferences { get; set; } = false;

        /// <summary>
        /// 是否跳过验证错误的行，默认为false
        /// </summary>
        public bool SkipValidationErrors { get; set; } = false;

        /// <summary>
        /// 是否跳过重复的记录，默认为false
        /// </summary>
        public bool SkipDuplicates { get; set; } = false;

        /// <summary>
        /// 批处理大小，默认为1000
        /// </summary>
        public int BatchSize { get; set; } = 1000;

        /// <summary>
        /// 是否启用进度报告，默认为true
        /// </summary>
        public bool EnableProgressReporting { get; set; } = true;

        /// <summary>
        /// 进度报告回调
        /// </summary>
        public Action<ImportProgress> ProgressCallback { get; set; }

        /// <summary>
        /// CSV分隔符，默认为逗号
        /// </summary>
        public char CsvSeparator { get; set; } = ',';

        /// <summary>
        /// 是否有表头行，默认为true
        /// </summary>
        public bool HasHeaderRow { get; set; } = true;
        
        /// <summary>
        /// 是否启用并行处理（适用于大数据量），默认为false
        /// </summary>
        public bool EnableParallelProcessing { get; set; } = false;
        
        /// <summary>
        /// 最大并行度（当EnableParallelProcessing为true时生效）
        /// </summary>
        public int MaxDegreeOfParallelism { get; set; } = Environment.ProcessorCount;
        
        /// <summary>
        /// 内存优化模式，适用于超大文件
        /// </summary>
        public bool MemoryOptimizedMode { get; set; } = false;
    }

    /// <summary>
    /// 导入进度信息
    /// </summary>
    public class ImportProgress
    {
        /// <summary>
        /// 当前处理的行数
        /// </summary>
        public int CurrentRow { get; set; }

        /// <summary>
        /// 总行数
        /// </summary>
        public int TotalRows { get; set; }

        /// <summary>
        /// 进度百分比（0-100）
        /// </summary>
        public int PercentageComplete => TotalRows > 0 ? (int)((double)CurrentRow / TotalRows * 100) : 0;

        /// <summary>
        /// 当前阶段描述
        /// </summary>
        public string Stage { get; set; }

        /// <summary>
        /// 成功处理的记录数
        /// </summary>
        public int SuccessfulRecords { get; set; }

        /// <summary>
        /// 失败的记录数
        /// </summary>
        public int FailedRecords { get; set; }
    }

    /// <summary>
    /// 导入字段信息
    /// </summary>
    public class ImportFieldInfo
    {
        /// <summary>
        /// 属性名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// 列名（显示名称）
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// 属性类型
        /// </summary>
        public Type PropertyType { get; set; }

        /// <summary>
        /// 是否必填
        /// </summary>
        public bool IsRequired { get; set; }

        /// <summary>
        /// 默认值
        /// </summary>
        public string DefaultValue { get; set; }

        /// <summary>
        /// 自定义格式
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 关联对象查找键
        /// </summary>
        public string LookupKey { get; set; }

        /// <summary>
        /// 是否有关联对象查找
        /// </summary>
        public bool HasLookup => !string.IsNullOrEmpty(LookupKey);

        /// <summary>
        /// 字段描述
        /// </summary>
        public string Description { get; set; }
    }
}