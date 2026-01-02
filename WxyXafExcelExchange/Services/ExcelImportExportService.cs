using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DevExpress.ExpressApp;
using DevExpress.Xpo;
using DevExpress.Data.Filtering;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Wxy.Xaf.ExcelExchange;
using Wxy.Xaf.ExcelExchange.Configuration;
using Wxy.Xaf.ExcelExchange.Services;
namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// Excel导入导出核心服务实现（平台无关）
    /// 统一使用UTF-8编码确保中文字符兼容性
    /// 集成新的双层特性系统支持
    /// </summary>
    public class ExcelImportExportService : IExcelImportExportService
    {
        private readonly ConfigurationManager _configurationManager;
        private readonly DataValidator _dataValidator;
        private readonly DataConverter _dataConverter;

        // 字典项缓存: key = "字典名:项名", value = DataDictionaryItem 对象
        private readonly Dictionary<string, object> _dataDictionaryItemCache = new Dictionary<string, object>();

        // DataDictionary 缓存: key = "字典名", value = DataDictionary 对象
        private readonly Dictionary<string, object> _dataDictionaryCache = new Dictionary<string, object>();

        /// <summary>
        /// 构造函数
        /// </summary>
        public ExcelImportExportService()
        {
            _configurationManager = new ConfigurationManager();
            _dataValidator = new DataValidator();
            _dataConverter = new DataConverter();
        }
        /// <summary>
        /// 导出数据到Excel/CSV
        /// </summary>
        public async Task<ExcelExportResult> ExportDataAsync(IEnumerable<object> data, Type objectType, ExcelExportOptions options = null)
        {
            var result = new ExcelExportResult();
            options = options ?? new ExcelExportOptions();
            try
            {
                // 获取Excel配置
                var config = _configurationManager.GetConfiguration(objectType);
                if (!config.IsExportEnabled)
                {
                    result.ErrorMessage = $"类型 {objectType.Name} 未启用Excel导出功能";
                    return result;
                }
                // 预加载集合属性（确保 XPCollection 数据被加载）
                
                foreach (var item in data)
                {
                    string parentKey = GetParentKeyValue(item);
                    foreach (var fieldConfig in config.FieldConfigurations)
                    {
                        bool isXPCollection = fieldConfig.PropertyInfo.PropertyType.Name.Contains("XPCollection") ||
                            (fieldConfig.PropertyInfo.PropertyType.FullName != null &&
                             fieldConfig.PropertyInfo.PropertyType.FullName.Contains("XPCollection"));
                        if (isXPCollection)
                        {
                            // 触发集合加载
                            var collection = fieldConfig.PropertyInfo.GetValue(item);
                            if (collection != null)
                            {
                                // 访问 Count 属性以触发加载
                                var countProp = collection.GetType().GetProperty("Count");
                                int count = (int)(countProp?.GetValue(collection) ?? 0);
                                // 检查是否配置为 MultiSheet
                                string exportFormat = fieldConfig.FieldAttribute.CollectionExportFormat;
                            }
                            else
                            {
                            }
                        }
                    }
                }
                // 确保统一使用UTF-8编码
                if (options.Encoding != Encoding.UTF8)
                {
                    options.Encoding = Encoding.UTF8;
                }
                var dataList = data?.ToList() ?? new List<object>();
                result.RecordCount = dataList.Count;
                if (dataList.Count == 0)
                {
                    result.ErrorMessage = "无数据可导出";
                    return result;
                }
                // 获取导出字段配置
                var exportFields = config.ExportFields;
                if (exportFields.Count == 0)
                {
                    result.ErrorMessage = "未找到可导出的字段";
                    return result;
                }
                // 应用类级别导出配置
                ApplyClassExportOptions(options, config.ClassConfiguration);
                // 根据格式生成内容
                switch (options.Format)
                {
                    case ExcelFormat.Csv:
                        result.FileContent = await GenerateCsvContentAsync(dataList, exportFields, options);
                        result.MimeType = "text/csv";
                        result.SuggestedFileName = $"{options.FileName ?? config.ClassConfiguration?.GetEffectiveDefaultFileName(objectType) ?? objectType.Name}_{DateTime.Now:yyyyMMddHHmmss}.csv";
                        break;
                    case ExcelFormat.Xlsx:
                        result.FileContent = await GenerateXlsxContentAsync(dataList, exportFields, options);
                        result.MimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                        result.SuggestedFileName = $"{options.FileName ?? config.ClassConfiguration?.GetEffectiveDefaultFileName(objectType) ?? objectType.Name}_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                        break;
                    default:
                        throw new NotSupportedException($"不支持的导出格式: {options.Format}");
                }
                result.IsSuccess = true;
                return result;
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.ErrorMessage = $"导出失败: {ex.Message}";
                return result;
            }
        }
        /// <summary>
        /// 从Excel/CSV导入数据（支持多 Sheet 子表导入）
        /// </summary>
        public async Task<ExcelImportResult> ImportDataAsync(byte[] fileContent, Type objectType, IObjectSpace objectSpace, ExcelImportOptions options = null, string fileName = null)
        {
            var result = new ExcelImportResult();
            options = options ?? new ExcelImportOptions();
            try
            {
                // **清空缓存** (每次导入都是新的开始)
                _dataDictionaryItemCache.Clear();
                _dataDictionaryCache.Clear();

                // **ReplaceAll 模式: 先删除所有现有记录**
                if (options.Mode == ExcelImportMode.ReplaceAll)
                {
                    try
                    {
                        // 获取所有现有对象
                        var existingObjects = objectSpace.GetObjects(objectType);
                        var existingList = existingObjects.Cast<object>().ToList();

                        if (existingList.Count > 0)
                        {
                            // 删除所有现有对象
                            foreach (var obj in existingList)
                            {
                                objectSpace.Delete(obj);
                            }

                            // 提交删除操作
                            objectSpace.CommitChanges();

                            // **修复**: ReplaceAll删除旧记录是正常操作，不应显示为警告
                            // 只在调试日志中记录
                            System.Diagnostics.Debug.WriteLine($"[ReplaceAll] 已删除 {existingList.Count} 条现有记录");
                        }
                    }
                    catch (Exception ex)
                    {
                        result.IsSuccess = false;
                        result.ErrorMessage = $"删除现有记录失败: {ex.Message}";
                        return result;
                    }
                }

                // 获取Excel配置
                var config = _configurationManager.GetConfiguration(objectType);
                if (!config.IsImportEnabled)
                {
                    result.ErrorMessage = $"类型 {objectType.Name} 未启用Excel导入功能";
                    return result;
                }
                // 应用类级别导入配置
                ApplyClassImportOptions(options, config.ClassConfiguration);
                // 检测是否为多 Sheet Excel 文件
                var isMultiSheetExcel = IsXlsxFile(fileContent) && HasMultipleSheets(fileContent);
                Dictionary<string, List<Dictionary<string, string>>> detailSheetData = null;
                if (isMultiSheetExcel)
                {
                    // 解析所有 Sheet
                    var allSheets = await Task.Run(() => ParseAllSheets(fileContent, options));
                    if (allSheets.Count == 0)
                    {
                        result.ErrorMessage = "文件中没有找到有效数据";
                        return result;
                    }
                    // 第一个 Sheet 作为主表数据
                    var mainSheet = allSheets.FirstOrDefault();
                    if (mainSheet == null)
                    {
                        result.ErrorMessage = "未找到主表数据";
                        return result;
                    }
                    // 检查第一个 Sheet 的列是否与当前对象类型匹配
                    // 如果不匹配，可能用户试图将父对象的多Sheet导出导入到子对象
                    var firstRowColumns = mainSheet.DataRows.FirstOrDefault()?.Keys.ToList() ?? new List<string>();
                    var importFields = config.ImportFields.Select(f => f.EffectiveColumnName).ToList();
                    // 检查列匹配度
                    int matchedColumns = 0;
                    foreach (var col in firstRowColumns)
                    {
                        if (importFields.Any(f => string.Equals(f, col, StringComparison.OrdinalIgnoreCase)))
                        {
                            matchedColumns++;
                        }
                    }
                    // 如果匹配的列数少于总列数的50%，认为对象类型不匹配
                    if (firstRowColumns.Count > 0 && matchedColumns < firstRowColumns.Count / 2)
                    {
                        result.ErrorMessage = $"导入文件与 {objectType.Name} 不匹配。" +
                            $"此文件可能是从其他对象（如父对象）导出的多Sheet文件。\n\n" +
                            $"提示: 多Sheet导入应该从主表（如 '订单'）进行导入，系统会自动处理所有明细表。";
                        result.IsSuccess = false;
                        return result;
                    }
                    // 其他 Sheet 可能是明细表数据
                    detailSheetData = allSheets.Skip(1).ToDictionary(s => s.SheetName, s => s.DataRows);
                    // 处理主表数据
                    result.TotalRecords = mainSheet.DataRows.Count;
                    var importResult = await ProcessMainTableImportAsync(mainSheet.DataRows, objectType, objectSpace, options, config, detailSheetData);
                    // 合并结果
                    result.SuccessCount = importResult.SuccessCount;
                    result.FailureCount = importResult.FailureCount;
                    result.Errors.AddRange(importResult.Errors);
                    result.Warnings.AddRange(importResult.Warnings);
                    result.IsSuccess = result.Errors.Count == 0 || result.SuccessCount > 0;
                }
                else
                {
                    // 单 Sheet 或 CSV 文件导入（原有逻辑）
                    var parseResult = await ParseFileContentAsync(fileContent, objectType, options, fileName);
                    if (parseResult == null || !parseResult.IsSuccess)
                    {
                        result.ErrorMessage = parseResult?.ErrorMessage ?? "解析文件失败";
                        return result;
                    }
                    result.TotalRecords = parseResult.DataRows?.Count ?? 0;
                    // 获取导入字段配置
                    var importFields = config.ImportFields;
                    var fieldConfigDict = importFields.ToDictionary(f => f.EffectiveColumnName, f => f);
                    // 处理数据行
                    int successCount = 0;
                    int failureCount = 0; // **新增**: 跟踪真正失败的记录数
                    int rowNumber = config.ClassConfiguration?.DataStartRowIndex ?? (options.HasHeaderRow ? 2 : 1);
                    // 🔴 调试输出：显示配置的字段
                    var debugLog = new StringBuilder();
                    debugLog.AppendLine($"=== 导入字段配置 ({objectType.Name}) ===");
                    foreach (var field in importFields.OrderBy(f => f.EffectiveSortOrder))
                    {
                        debugLog.AppendLine($"  [{field.EffectiveSortOrder}] {field.PropertyInfo.Name} -> Excel列名: '{field.EffectiveColumnName}', 必填: {field.FieldAttribute.IsRequired}, 启用导入: {field.FieldAttribute.EnabledImport}");
                    }
                    debugLog.AppendLine($"=== 总字段数: {importFields.Count} ===");
                    // 🔴 调试输出：显示Excel第一行的列
                    if (parseResult.DataRows != null && parseResult.DataRows.Count > 0)
                    {
                        var firstRow = parseResult.DataRows.First();
                        debugLog.AppendLine($"=== Excel第一行数据 (行{rowNumber}) ===");
                        foreach (var col in firstRow.Keys)
                        {
                            debugLog.AppendLine($"  Excel列: '{col}' = '{firstRow[col]}'");
                        }
                        debugLog.AppendLine($"=== Excel列数: {firstRow.Count} ===");
                    }
                    // 写入调试日志文件
                    try
                    {
                        var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "import_debug.log");
                        File.AppendAllText(logPath, debugLog.ToString(), Encoding.UTF8);
                    }
                    catch
                    {
                        // 写入日志文件失败,忽略但不影响导入流程
                        System.Diagnostics.Debug.WriteLine("[ExcelImport] Failed to write debug log file");
                    }
                    if (parseResult.DataRows != null)
                    {
                        foreach (var dataRow in parseResult.DataRows)
                    {
                        try
                        {
                            var importRowResult = await ProcessImportRowAsync(dataRow, objectType, objectSpace, fieldConfigDict, options, config, rowNumber);
                            if (importRowResult.IsSuccess)
                            {
                                // **修复**: 只有当真正创建/更新了对象时才计入成功
                                // 如果只是跳过记录(CreatedObject为null),不计入成功
                                if (importRowResult.CreatedObject != null)
                                {
                                    successCount++;
                                }
                            }
                            else
                            {
                                // **修复**: 只有处理过程真正失败时才计入失败
                                failureCount++;
                                result.Errors.AddRange(importRowResult.Errors);
                                // 检查验证模式
                                if (config.ClassConfiguration?.ValidationMode == Enums.ValidationMode.Strict && importRowResult.Errors.Count > 0)
                                {
                                    result.Errors.Add(new ExcelImportError
                                    {
                                        RowNumber = rowNumber,
                                        FieldName = "系统",
                                        ErrorMessage = "严格验证模式下遇到错误，停止导入",
                                        ErrorType = ExcelImportErrorType.SystemError
                                    });
                                    break;
                                }
                            }
                            result.Warnings.AddRange(importRowResult.Warnings);
                            // 检查错误数量限制
                            if (result.Errors.Count >= options.MaxErrors)
                            {
                                result.Errors.Add(new ExcelImportError
                                {
                                    RowNumber = rowNumber,
                                    FieldName = "系统",
                                    ErrorMessage = $"错误数量已达到最大限制 ({options.MaxErrors})，停止导入",
                                    ErrorType = ExcelImportErrorType.SystemError
                                });
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            result.Errors.Add(new ExcelImportError
                            {
                                RowNumber = rowNumber,
                                FieldName = "行处理",
                                ErrorMessage = ex.Message,
                                ErrorType = ExcelImportErrorType.SystemError
                            });
                        }
                        rowNumber++;
                    }
                    }
                    result.SuccessCount = successCount;
                    result.FailureCount = failureCount; // **修复**: 使用真正的失败计数，而不是计算值
                    result.IsSuccess = result.Errors.Count == 0 || result.SuccessCount > 0;
                }
                return result;
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.ErrorMessage = $"导入失败: {ex.Message}";
                result.Errors.Add(new ExcelImportError
                {
                    RowNumber = 0,
                    FieldName = "系统",
                    ErrorMessage = ex.Message,
                    ErrorType = ExcelImportErrorType.SystemError
                });
                return result;
            }
        }
        /// <summary>
        /// 验证文件格式
        /// </summary>
        public FileValidationResult ValidateFile(string fileName, byte[] fileContent)
        {
            var result = new FileValidationResult
            {
                FileSize = fileContent?.Length ?? 0
            };
            try
            {
                if (fileContent == null || fileContent.Length == 0)
                {
                    result.ErrorMessage = "文件内容为空";
                    return result;
                }
                // 检查文件大小（默认最大10MB）
                const long maxFileSize = 10 * 1024 * 1024;
                if (fileContent.Length > maxFileSize)
                {
                    result.ErrorMessage = $"文件大小超过限制 ({maxFileSize / 1024 / 1024}MB)";
                    return result;
                }
                // 根据文件扩展名和内容检测格式
                var extension = Path.GetExtension(fileName)?.ToLowerInvariant();
                
                switch (extension)
                {
                    case ".csv":
                        result.DetectedFormat = ExcelFormat.Csv;
                        result.IsValid = ValidateCsvContent(fileContent);
                        break;
                    case ".xlsx":
                        result.DetectedFormat = ExcelFormat.Xlsx;
                        result.IsValid = ValidateXlsxContent(fileContent);
                        break;
                    case ".xls":
                        result.DetectedFormat = ExcelFormat.Xls;
                        result.IsValid = ValidateXlsContent(fileContent);
                        break;
                    default:
                        // 尝试自动检测格式
                        if (ValidateCsvContent(fileContent))
                        {
                            result.DetectedFormat = ExcelFormat.Csv;
                            result.IsValid = true;
                        }
                        else if (ValidateXlsxContent(fileContent))
                        {
                            result.DetectedFormat = ExcelFormat.Xlsx;
                            result.IsValid = true;
                        }
                        else
                        {
                            result.ErrorMessage = "不支持的文件格式";
                        }
                        break;
                }
                if (!result.IsValid && string.IsNullOrEmpty(result.ErrorMessage))
                {
                    result.ErrorMessage = "文件格式验证失败";
                }
                return result;
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.ErrorMessage = $"文件验证失败: {ex.Message}";
                return result;
            }
        }
        /// <summary>
        /// 获取对象类型的可导入/导出属性
        /// </summary>
        public List<ExcelFieldInfo> GetExcelFields(Type objectType)
        {
            try
            {
                // 使用新的配置管理器获取配置
                var config = _configurationManager.GetConfiguration(objectType);
                
                // 转换为旧的ExcelFieldInfo格式以保持向后兼容
                var fields = new List<ExcelFieldInfo>();
                
                foreach (var fieldConfig in config.FieldConfigurations)
                {
                    var fieldInfo = new ExcelFieldInfo
                    {
                        PropertyName = fieldConfig.PropertyInfo.Name,
                        DisplayName = fieldConfig.EffectiveColumnName,
                        DataType = fieldConfig.PropertyInfo.PropertyType,
                        CanExport = fieldConfig.FieldAttribute.EnabledExport,
                        CanImport = fieldConfig.FieldAttribute.EnabledImport,
                        IsRequired = fieldConfig.FieldAttribute.IsRequired,
                        Order = fieldConfig.EffectiveSortOrder
                    };
                    
                    fields.Add(fieldInfo);
                }
                return fields.OrderBy(f => f.Order).ToList();
            }
            catch (Exception ex)
            {
                throw new Exception($"获取Excel字段信息失败: {ex.Message}", ex);
            }
        }
        #region 私有方法
        /// <summary>
        /// 检查是否为 XLSX 文件
        /// </summary>
        private bool IsXlsxFile(byte[] fileContent)
        {
            return fileContent.Length > 4 &&
                   fileContent[0] == 0x50 && fileContent[1] == 0x4B &&
                   fileContent[2] == 0x03 && fileContent[3] == 0x04;
        }
        /// <summary>
        /// 检查是否有多个 Sheet
        /// </summary>
        private bool HasMultipleSheets(byte[] fileContent)
        {
            try
            {
                using (var stream = new MemoryStream(fileContent))
                {
                    using (var spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
                    {
                        var sheets = spreadsheetDocument.WorkbookPart?.Workbook?.Sheets;
                        return sheets != null && sheets.Count() > 1;
                    }
                }
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 处理主表导入（包含子表数据处理）
        /// </summary>
        private async Task<ExcelImportResult> ProcessMainTableImportAsync(
            List<Dictionary<string, string>> mainDataRows,
            Type objectType,
            IObjectSpace objectSpace,
            ExcelImportOptions options,
            ExcelConfiguration config,
            Dictionary<string, List<Dictionary<string, string>>> detailSheetData)
        {
            var result = new ExcelImportResult();
            try
            {
                // 获取导入字段配置
                var importFields = config.ImportFields;
                var fieldConfigDict = importFields.ToDictionary(f => f.EffectiveColumnName, f => f);
                // 找出集合字段配置
                var collectionFields = config.FieldConfigurations
                    .Where(f => f.FieldAttribute.CollectionExportFormat == "MultiSheet")
                    .ToList();
                int successCount = 0;
                int failureCount = 0; // **新增**: 跟踪真正失败的记录数
                int rowNumber = config.ClassConfiguration?.DataStartRowIndex ?? (options.HasHeaderRow ? 2 : 1);
                // 存储已创建的主表对象（用于子表关联）
                var createdObjects = new Dictionary<string, object>();
                foreach (var dataRow in mainDataRows)
                {
                    try
                    {
                        var importRowResult = await ProcessImportRowAsync(dataRow, objectType, objectSpace, fieldConfigDict, options, config, rowNumber);
                        if (importRowResult.IsSuccess)
                        {
                            // **修复**: 只有当真正创建/更新了对象时才计入成功
                            // 如果只是跳过记录(CreatedObject为null),不计入成功
                            if (importRowResult.CreatedObject != null)
                            {
                                System.Diagnostics.Debug.WriteLine($"[统计] 行{rowNumber}: CreatedObject不为null, successCount增加到{successCount + 1}");
                                successCount++;
                                var parentKey = GetParentKeyValueFromRow(dataRow, config);
                                createdObjects[parentKey] = importRowResult.CreatedObject;
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"[统计] 行{rowNumber}: CreatedObject为null,不计入successCount (当前successCount={successCount})");
                                System.Diagnostics.Debug.WriteLine($"[统计] 行{rowNumber}: 警告数量={importRowResult.Warnings.Count}");
                                if (importRowResult.Warnings.Count > 0)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[统计] 行{rowNumber}: 第一个警告: {importRowResult.Warnings[0].WarningMessage}");
                                }
                            }
                            // 即使跳过记录,也要收集警告
                            result.Warnings.AddRange(importRowResult.Warnings);
                        }
                        else
                        {
                            // **修复**: 只有处理过程真正失败时才计入失败
                            failureCount++;
                            result.Errors.AddRange(importRowResult.Errors);
                            if (config.ClassConfiguration?.ValidationMode == Enums.ValidationMode.Strict && importRowResult.Errors.Count > 0)
                            {
                                result.Errors.Add(new ExcelImportError
                                {
                                    RowNumber = rowNumber,
                                    FieldName = "系统",
                                    ErrorMessage = "严格验证模式下遇到错误，停止导入",
                                    ErrorType = ExcelImportErrorType.SystemError
                                });
                                break;
                            }
                            result.Warnings.AddRange(importRowResult.Warnings);
                        }
                        if (result.Errors.Count >= options.MaxErrors)
                        {
                            result.Errors.Add(new ExcelImportError
                            {
                                RowNumber = rowNumber,
                                FieldName = "系统",
                                ErrorMessage = $"错误数量已达到最大限制 ({options.MaxErrors})，停止导入",
                                ErrorType = ExcelImportErrorType.SystemError
                            });
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Errors.Add(new ExcelImportError
                        {
                            RowNumber = rowNumber,
                            FieldName = "行处理",
                            ErrorMessage = ex.Message,
                            ErrorType = ExcelImportErrorType.SystemError
                        });
                    }
                    rowNumber++;
                }
                // 处理子表数据
                if (detailSheetData != null && detailSheetData.Count > 0 && collectionFields.Count > 0)
                {
                    await ProcessDetailSheetsAsync(detailSheetData, collectionFields, createdObjects, objectSpace, config, result, objectType, options);
                }
                result.SuccessCount = successCount;
                result.FailureCount = failureCount; // **修复**: 使用真正的失败计数，而不是计算值
                result.IsSuccess = result.Errors.Count == 0 || result.SuccessCount > 0;
                System.Diagnostics.Debug.WriteLine($"[统计] ===== 导入完成 =====");
                System.Diagnostics.Debug.WriteLine($"[统计] 总记录数: {mainDataRows.Count}");
                System.Diagnostics.Debug.WriteLine($"[统计] 成功导入: {successCount}");
                System.Diagnostics.Debug.WriteLine($"[统计] 失败/跳过: {result.FailureCount}");
                System.Diagnostics.Debug.WriteLine($"[统计] 警告数量: {result.Warnings.Count}");
                return result;
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.ErrorMessage = $"主表导入处理失败: {ex.Message}";
                return result;
            }
        }
        /// <summary>
        /// 处理子表 Sheet 数据
        /// </summary>
        private async Task ProcessDetailSheetsAsync(
            Dictionary<string, List<Dictionary<string, string>>> detailSheetData,
            List<ExcelFieldConfiguration> collectionFields,
            Dictionary<string, object> parentObjects,
            IObjectSpace objectSpace,
            ExcelConfiguration config,
            ExcelImportResult result,
            Type objectType,
            ExcelImportOptions options)
        {
            foreach (var collectionField in collectionFields)
            {
                // 查找对应的明细 Sheet
                string detailSheetName = string.IsNullOrEmpty(collectionField.FieldAttribute.DetailSheetName)
                    ? collectionField.EffectiveColumnName + "明细"
                    : collectionField.FieldAttribute.DetailSheetName;
                if (!detailSheetData.ContainsKey(detailSheetName))
                {
                    continue;
                }
                var detailRows = detailSheetData[detailSheetName];
                // 获取明细对象类型
                var collectionType = collectionField.PropertyInfo.PropertyType;
                Type detailType = null;
                if (collectionType.IsGenericType)
                {
                    detailType = collectionType.GetGenericArguments()[0];
                }
                if (detailType == null)
                {
                    result.Errors.Add(new ExcelImportError
                    {
                        RowNumber = 0,
                        FieldName = collectionField.PropertyInfo.Name,
                        ErrorMessage = "无法确定明细对象类型",
                        ErrorType = ExcelImportErrorType.SystemError
                    });
                    continue;
                }
                // 获取明细对象的导入配置
                var detailConfig = _configurationManager.GetConfiguration(detailType);
                if (!detailConfig.IsImportEnabled)
                {
                    result.Errors.Add(new ExcelImportError
                    {
                        RowNumber = 0,
                        FieldName = collectionField.PropertyInfo.Name,
                        ErrorMessage = $"明细对象 {detailType.Name} 未启用导入功能",
                        ErrorType = ExcelImportErrorType.SystemError
                    });
                    continue;
                }
                var detailImportFields = detailConfig.ImportFields;
                var detailFieldDict = detailImportFields.ToDictionary(f => f.EffectiveColumnName, f => f);
                // 获取关联字段名称
                string configuredRelationFieldName = collectionField.FieldAttribute.RelationFieldName ?? "关联主表记录";
                string actualRelationFieldName = null;
                // 尝试在明细数据的第一行中查找实际的关联字段列名
                if (detailRows.Count > 0)
                {
                    var firstRow = detailRows[0];
                    // 优先使用配置的关联字段名
                    if (firstRow.ContainsKey(configuredRelationFieldName))
                    {
                        actualRelationFieldName = configuredRelationFieldName;
                    }
                    else
                    {
                        // 尝试自动检测关联字段（通过父对象的主键属性名）
                        string parentKeyName = GetParentKeyPropertyName(config.ObjectType);
                        // 在明细表的列中查找包含父对象主键属性名的列
                        foreach (var colName in firstRow.Keys)
                        {
                            if (string.Equals(colName, parentKeyName, StringComparison.OrdinalIgnoreCase) ||
                                colName.Contains(parentKeyName) ||
                                string.Equals(colName, configuredRelationFieldName, StringComparison.OrdinalIgnoreCase))
                            {
                                actualRelationFieldName = colName;
                                break;
                            }
                        }
                        // 如果还没找到，尝试查找常见的关联字段名
                        if (string.IsNullOrEmpty(actualRelationFieldName))
                        {
                            var commonRelationNames = new[] { "订单编号", "OrderNo", "ParentId", "ParentKey", "关联主表记录" };
                            foreach (var name in commonRelationNames)
                            {
                                if (firstRow.ContainsKey(name))
                                {
                                    actualRelationFieldName = name;
                                    break;
                                }
                            }
                        }
                        if (string.IsNullOrEmpty(actualRelationFieldName))
                        {
                            result.Warnings.Add(new ExcelImportWarning
                            {
                                RowNumber = 0,
                                FieldName = "关联字段",
                                WarningMessage = $"未找到关联字段列（配置: {configuredRelationFieldName}），明细数据无法关联到主表"
                            });
                            continue;
                        }
                    }
                }
                // 处理每一行明细数据
                foreach (var detailRow in detailRows)
                {
                    try
                    {
                        // 检查是否包含关联字段
                        if (!detailRow.ContainsKey(actualRelationFieldName) || string.IsNullOrEmpty(detailRow[actualRelationFieldName]))
                        {
                            result.Warnings.Add(new ExcelImportWarning
                            {
                                RowNumber = 0,
                                FieldName = actualRelationFieldName,
                                WarningMessage = $"明细数据缺少关联字段值，跳过此行"
                            });
                            continue;
                        }
                        string parentKey = detailRow[actualRelationFieldName];
                        // 查找父对象
                        if (!parentObjects.ContainsKey(parentKey))
                        {
                            result.Warnings.Add(new ExcelImportWarning
                            {
                                RowNumber = 0,
                                FieldName = actualRelationFieldName,
                                WarningMessage = $"未找到关联的主表记录: {parentKey}"
                            });
                            continue;
                        }
                        var parentObject = parentObjects[parentKey];
                        // 查找或创建明细对象（支持重复检查）
                        object detailObject = null;
                        // 根据导入模式决定是否查找现有对象
                        if (options.Mode == ExcelImportMode.CreateOrUpdate || options.Mode == ExcelImportMode.UpdateOnly)
                        {
                            // 尝试查找现有明细对象
                            detailObject = FindExistingDetailObject(objectSpace, detailType, detailRow, detailFieldDict, parentObject);
                            if (detailObject != null)
                            {
                            }
                        }
                        // 如果没找到现有对象，则创建新对象
                        if (detailObject == null)
                        {
                            // 只有在 CreateOnly 或 CreateOrUpdate 模式下才创建新对象
                            if (options.Mode == ExcelImportMode.CreateOnly || options.Mode == ExcelImportMode.CreateOrUpdate)
                            {
                                detailObject = objectSpace.CreateObject(detailType);
                            }
                            else
                            {
                                // UpdateOnly 模式下未找到对象则跳过
                                result.Warnings.Add(new ExcelImportWarning
                                {
                                    RowNumber = 0,
                                    FieldName = "明细对象",
                                    WarningMessage = $"未找到要更新的明细对象，跳过"
                                });
                                continue;
                            }
                        }
                        else
                        {
                        }
                        // 设置明细对象属性
                        foreach (var kvp in detailRow)
                        {
                            if (kvp.Key == actualRelationFieldName) continue; // 跳过关联字段
                            if (detailFieldDict.TryGetValue(kvp.Key, out var fieldConfig))
                            {
                                try
                                {
                                    var property = fieldConfig.PropertyInfo;
                                    if (property != null && property.CanWrite)
                                    {
                                        object convertedValue = null;
                                        // 检查是否为XPO关联对象
                                        bool isXpoObject = property.PropertyType.GetInterface("DevExpress.Xpo.IXPObject") != null;
                                        if (isXpoObject)
                                        {
                                            // XPO关联对象处理
                                            string matchFieldName = fieldConfig.FieldAttribute.ReferenceMatchField ?? "Oid";
                                            // 使用ObjectSpace查询关联对象
                                            var referenceType = property.PropertyType;
                                            var criteria = CriteriaOperator.Parse($"[{matchFieldName}] = ?", kvp.Value);
                                            // IObjectSpace 不直接支持 CriteriaOperator，需要使用 GetObjects
                                            var referencedObjects = objectSpace.GetObjects(referenceType, criteria);
                                            var referencedObject = referencedObjects.Cast<object>().FirstOrDefault();
                                            if (referencedObject != null)
                                            {
                                                convertedValue = referencedObject;
                                            }
                                            else if (fieldConfig.FieldAttribute.ReferenceCreateIfNotExists)
                                            {
                                                // 自动创建关联对象
                                                convertedValue = objectSpace.CreateObject(referenceType);
                                                // 设置匹配字段
                                                var matchProperty = referenceType.GetProperty(matchFieldName);
                                                if (matchProperty != null && matchProperty.CanWrite)
                                                {
                                                    var matchConvertResult = _dataConverter.ConvertFromExcel(kvp.Value, new ExcelFieldConfiguration { PropertyInfo = matchProperty, FieldAttribute = new ExcelFieldAttribute() });
                                                    if (matchConvertResult.IsSuccess)
                                                    {
                                                        matchProperty.SetValue(convertedValue, matchConvertResult.ConvertedValue);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                result.Warnings.Add(new ExcelImportWarning
                                                {
                                                    RowNumber = 0,
                                                    FieldName = kvp.Key,
                                                    WarningMessage = $"未找到关联的{referenceType.Name}对象: {kvp.Value}"
                                                });
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            // 普通类型转换
                                            var convertResult = _dataConverter.ConvertFromExcel(kvp.Value, fieldConfig);
                                            if (!convertResult.IsSuccess)
                                            {
                                                result.Warnings.Add(new ExcelImportWarning
                                                {
                                                    RowNumber = 0,
                                                    FieldName = kvp.Key,
                                                    WarningMessage = convertResult.ErrorMessage
                                                });
                                                continue;
                                            }
                                            convertedValue = convertResult.ConvertedValue;
                                        }
                                        // 设置属性值
                                        property.SetValue(detailObject, convertedValue);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    result.Warnings.Add(new ExcelImportWarning
                                    {
                                        RowNumber = 0,
                                        FieldName = kvp.Key,
                                        WarningMessage = $"字段设置失败: {ex.Message}"
                                    });
                                }
                            }
                        }
                        // 设置关联到父对象
                        // 查找明细类型中指向父对象的属性（类型为主表类型）
                        PropertyInfo detailParentProperty = null;
                        foreach (var prop in detailType.GetProperties())
                        {
                            if (prop.PropertyType == objectType && prop.CanWrite)
                            {
                                detailParentProperty = prop;
                                break;
                            }
                        }
                        if (detailParentProperty != null)
                        {
                            detailParentProperty.SetValue(detailObject, parentObject);
                        }
                        else
                        {
                            // 备用方法：通过父对象的集合添加
                            var parentCollection = collectionField.PropertyInfo.GetValue(parentObject);
                            if (parentCollection != null)
                            {
                                var addMethod = parentCollection.GetType().GetMethod("Add");
                                addMethod?.Invoke(parentCollection, new object[] { detailObject });
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Errors.Add(new ExcelImportError
                        {
                            RowNumber = 0,
                            FieldName = "明细数据处理",
                            ErrorMessage = ex.Message,
                            ErrorType = ExcelImportErrorType.SystemError
                        });
                    }
                }
            }
        }
        /// <summary>
        /// 从数据行获取主表的关联键值
        /// </summary>
        private string GetParentKeyValueFromRow(Dictionary<string, string> dataRow, ExcelConfiguration config)
        {
            if (dataRow == null || dataRow.Count == 0)
            {
                return "";
            }
            // 获取对象类型
            var objectType = config?.ObjectType;
            if (objectType == null)
            {
                return dataRow.Values.FirstOrDefault(v => !string.IsNullOrEmpty(v)) ?? "";
            }
            // 尝试获取具有 DefaultProperty 的属性名称
            var defaultPropertyAttr = objectType.GetCustomAttributes(typeof(System.ComponentModel.DefaultPropertyAttribute), false)
                .FirstOrDefault() as System.ComponentModel.DefaultPropertyAttribute;
            string keyPropertyName = null;
            if (defaultPropertyAttr != null && !string.IsNullOrEmpty(defaultPropertyAttr.Name))
            {
                keyPropertyName = defaultPropertyAttr.Name;
            }
            else
            {
                // 尝试获取名为 "Name"、"Title"、"Code" 或 "订单编号" 的属性
                keyPropertyName = "Name";
                if (objectType.GetProperty("Name") == null)
                {
                    keyPropertyName = "Title";
                    if (objectType.GetProperty("Title") == null)
                    {
                        keyPropertyName = "Code";
                        if (objectType.GetProperty("Code") == null)
                        {
                            keyPropertyName = "订单编号";
                        }
                    }
                }
            }
            // 在配置中查找对应的列名
            var keyFieldConfig = config?.FieldConfigurations?.FirstOrDefault(f => f.PropertyInfo.Name == keyPropertyName);
            string keyColumnName = keyFieldConfig?.EffectiveColumnName ?? keyPropertyName;
            // 从数据行中获取键值
            if (dataRow.TryGetValue(keyColumnName, out var keyValue))
            {
                if (!string.IsNullOrEmpty(keyValue))
                {
                    return keyValue;
                }
            }
            // 如果没有找到，尝试忽略大小写匹配
            foreach (var kvp in dataRow)
            {
                if (string.Equals(kvp.Key, keyColumnName, StringComparison.OrdinalIgnoreCase))
                {
                    if (!string.IsNullOrEmpty(kvp.Value))
                    {
                        return kvp.Value;
                    }
                }
            }
            // 降级：返回第一个非空值
            foreach (var kvp in dataRow)
            {
                if (!string.IsNullOrEmpty(kvp.Value))
                {
                    return kvp.Value;
                }
            }
            return "";
        }
        /// <summary>
        /// 获取父对象的主键属性名称（用于在子表中查找关联字段）
        /// </summary>
        private string GetParentKeyPropertyName(Type parentType)
        {
            if (parentType == null)
            {
                return "订单编号";
            }
            // 尝试获取具有 DefaultProperty 的属性名称
            var defaultPropertyAttr = parentType.GetCustomAttributes(typeof(System.ComponentModel.DefaultPropertyAttribute), false)
                .FirstOrDefault() as System.ComponentModel.DefaultPropertyAttribute;
            if (defaultPropertyAttr != null && !string.IsNullOrEmpty(defaultPropertyAttr.Name))
            {
                return defaultPropertyAttr.Name;
            }
            // 尝试获取名为 "Name"、"Title"、"Code" 或 "订单编号" 的属性
            if (parentType.GetProperty("Name") != null)
            {
                return "Name";
            }
            if (parentType.GetProperty("Title") != null)
            {
                return "Title";
            }
            if (parentType.GetProperty("Code") != null)
            {
                return "Code";
            }
            if (parentType.GetProperty("订单编号") != null)
            {
                return "订单编号";
            }
            // 默认返回"订单编号"
            return "订单编号";
        }
        /// <summary>
        /// 应用类级别导入配置
        /// </summary>
        /// <param name="options">导入选项</param>
        /// <param name="classConfig">类级别配置</param>
        private void ApplyClassImportOptions(ExcelImportOptions options, ExcelImportExportAttribute classConfig)
        {
            if (classConfig == null) return;

            // **修复**: 如果用户已经显式指定了导入模式,则不应用类配置的默认模式
            // 这样可以避免覆盖用户在对话框中选择的行为
            if (options.IsUserSpecifiedMode)
            {
                System.Diagnostics.Debug.WriteLine($"[ApplyClassImportOptions] 用户已显式指定模式 {options.Mode},跳过类配置的应用");
                return;
            }

            // 应用重复策略(仅作为默认值)
            switch (classConfig.ImportDuplicateStrategy)
            {
                case Enums.ImportDuplicateStrategy.Insert:
                    options.Mode = ExcelImportMode.CreateOnly;
                    break;
                case Enums.ImportDuplicateStrategy.Update:
                    options.Mode = ExcelImportMode.UpdateOnly;
                    break;
                case Enums.ImportDuplicateStrategy.InsertOrUpdate:
                    options.Mode = ExcelImportMode.CreateOrUpdate;
                    break;
                case Enums.ImportDuplicateStrategy.Ignore:
                    options.SkipDuplicates = true;
                    break;
            }
        }
        /// <summary>
        /// 应用类级别导出配置
        /// </summary>
        /// <param name="options">导出选项</param>
        /// <param name="classConfig">类级别配置</param>
        private void ApplyClassExportOptions(ExcelExportOptions options, ExcelImportExportAttribute classConfig)
        {
            if (classConfig == null) return;
            // 应用表头设置
            if (!classConfig.ExportIncludeHeader)
            {
                options.IncludeHeaders = false;
            }
            // 其他类级别配置可以在这里应用
        }
        /// <summary>
        /// 生成CSV内容
        /// </summary>
        private async Task<byte[]> GenerateCsvContentAsync(List<object> data, List<ExcelFieldConfiguration> fields, ExcelExportOptions options)
        {
            return await Task.Run(() =>
            {
                var sb = new StringBuilder();
                // 写入表头
                if (options.IncludeHeaders)
                {
                    for (int i = 0; i < fields.Count; i++)
                    {
                        if (i > 0) sb.Append(",");
                        sb.Append($"\"{fields[i].EffectiveColumnName}\"");
                    }
                    sb.AppendLine();
                }
                // 写入数据
                foreach (var item in data)
                {
                    for (int i = 0; i < fields.Count; i++)
                    {
                        if (i > 0) sb.Append(",");
                        try
                        {
                            var fieldConfig = fields[i];
                            var property = item.GetType().GetProperty(fieldConfig.PropertyInfo.Name);
                            var value = property?.GetValue(item);
                            
                            // 使用数据转换器格式化值
                            var convertResult = _dataConverter.ConvertToExcel(value, fieldConfig);
                            string cellValue = convertResult.IsSuccess ? convertResult.ConvertedValue?.ToString() ?? "" : value?.ToString() ?? "";
                            // CSV转义处理
                            if (cellValue.Contains(",") || cellValue.Contains("\n") || cellValue.Contains("\""))
                            {
                                cellValue = $"\"{cellValue.Replace("\"", "\"\"")}\"";
                            }
                            sb.Append(cellValue);
                        }
                        catch
                        {
                            sb.Append("");
                        }
                    }
                    sb.AppendLine();
                }
                // 转换为字节数组，确保中文字符正确编码
                var content = sb.ToString();
                
                // 检查内容是否包含中文字符
                bool hasChineseContent = EncodingDetector.ContainsChineseCharacters(content);
                if (hasChineseContent)
                {
                }
                if (options.AddBom)
                {
                    var bom = options.Encoding.GetPreamble();
                    var contentBytes = options.Encoding.GetBytes(content);
                    var result = new byte[bom.Length + contentBytes.Length];
                    Buffer.BlockCopy(bom, 0, result, 0, bom.Length);
                    Buffer.BlockCopy(contentBytes, 0, result, bom.Length, contentBytes.Length);
                    return result;
                }
                else
                {
                    return options.Encoding.GetBytes(content);
                }
            });
        }
        /// <summary>
        /// 生成 XLSX 内容（支持多 Sheet）
        /// </summary>
        private async Task<byte[]> GenerateXlsxContentAsync(List<object> data, List<ExcelFieldConfiguration> fields, ExcelExportOptions options)
        {
            return await Task.Run(() =>
            {
                using (var stream = new MemoryStream())
                {
                    using (var spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
                    {
                        var workbookPart = spreadsheetDocument.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();
                        var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                        // 检测 XPCollection 字段并筛选出需要作为多 Sheet 导出的集合属性
                        var collectionFields = fields.Where(f =>
                            f.PropertyInfo.PropertyType.Name.Contains("XPCollection") ||
                            f.PropertyInfo.PropertyType.Name.StartsWith("XPCollection") ||
                            (f.PropertyInfo.PropertyType.FullName != null && f.PropertyInfo.PropertyType.FullName.Contains("XPCollection")) ||
                            (f.PropertyInfo.PropertyType.IsGenericType &&
                             f.PropertyInfo.PropertyType.GetGenericTypeDefinition().Name.Contains("XPCollection")))
                            .ToList();
                        // 筛选出配置为 MultiSheet 模式的集合字段
                        var multiSheetCollectionFields = collectionFields
                            .Where(f => f.FieldAttribute.CollectionExportFormat == "MultiSheet")
                            .ToList();
                        foreach (var cf in multiSheetCollectionFields)
                        {
                        }
                        // 生成主数据 Sheet（排除所有集合字段）
                        string mainSheetName = options.SheetName ?? "主表";
                        CreateWorksheet(spreadsheetDocument, sheets, mainSheetName, data, fields, excludeCollectionFields: true);
                        // 生成子对象 MultiSheet（如果有）
                        foreach (var collectionField in multiSheetCollectionFields)
                        {
                            // 使用自定义的 Sheet 名称或默认名称
                            string detailSheetName = string.IsNullOrEmpty(collectionField.FieldAttribute.DetailSheetName)
                                ? collectionField.EffectiveColumnName + "明细"
                                : collectionField.FieldAttribute.DetailSheetName;
                            CreateDetailWorksheet(spreadsheetDocument, sheets, detailSheetName, data, collectionField);
                        }
                        spreadsheetDocument.WorkbookPart.Workbook.Save();
                    }
                    return stream.ToArray();
                }
            });
        }
        /// <summary>
        /// 创建主数据工作表
        /// </summary>
        private void CreateWorksheet(SpreadsheetDocument spreadsheet, Sheets sheets, string sheetName,
            List<object> data, List<ExcelFieldConfiguration> fields, bool excludeCollectionFields = false)
        {
            var worksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            var sheet = new Sheet()
            {
                Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = (uint)(sheets.Count() + 1),
                Name = sheetName
            };
            sheets.Append(sheet);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            // 过滤字段
            var exportFields = excludeCollectionFields
                ? fields.Where(f => !f.PropertyInfo.PropertyType.Name.Contains("XPCollection") &&
                    (f.PropertyInfo.PropertyType.FullName == null || !f.PropertyInfo.PropertyType.FullName.Contains("XPCollection")) &&
                    (!f.PropertyInfo.PropertyType.IsGenericType ||
                     !f.PropertyInfo.PropertyType.GetGenericTypeDefinition().Name.Contains("XPCollection")))
                    .ToList()
                : fields;
            // 添加表头
            var headerRow = new Row();
            uint columnIndex = 1;
            foreach (var field in exportFields)
            {
                var cell = CreateTextCell(columnIndex++, 1, field.EffectiveColumnName);
                headerRow.AppendChild(cell);
            }
            sheetData.AppendChild(headerRow);
            // 添加数据行
            uint rowIndex = 2;
            foreach (var item in data)
            {
                var dataRow = new Row() { RowIndex = rowIndex++ };
                columnIndex = 1;
                foreach (var field in exportFields)
                {
                    try
                    {
                        var propValue = field.PropertyInfo.GetValue(item);
                        var convertResult = _dataConverter.ConvertToExcel(propValue, field);
                        string cellValue = convertResult?.ConvertedValue?.ToString() ?? "";
                        var cell = CreateTextCell(columnIndex++, rowIndex - 1, cellValue);
                        dataRow.AppendChild(cell);
                    }
                    catch
                    {
                        var cell = CreateTextCell(columnIndex++, rowIndex - 1, "");
                        dataRow.AppendChild(cell);
                    }
                }
                sheetData.AppendChild(dataRow);
            }
        }
        /// <summary>
        /// 创建明细工作表（从集合字段展开）
        /// </summary>
        private void CreateDetailWorksheet(SpreadsheetDocument spreadsheet, Sheets sheets, string sheetName,
            List<object> data, ExcelFieldConfiguration collectionField)
        {
            var worksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            var sheet = new Sheet()
            {
                Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = (uint)(sheets.Count() + 1),
                Name = sheetName
            };
            sheets.Append(sheet);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            // 获取集合项的类型
            var collectionType = collectionField.PropertyInfo.PropertyType;
            Type itemType = null;
            if (collectionType.IsGenericType)
            {
                itemType = collectionType.GetGenericArguments()[0];
            }
            if (itemType == null)
            {
                return; // 无法确定集合项类型
            }
            // 获取集合项的可导出字段
            var itemConfig = _configurationManager.GetConfiguration(itemType);
            var itemFields = itemConfig.ExportFields;
            // 获取关联字段名称（用于关联主表）
            string relationFieldName = collectionField.FieldAttribute.RelationFieldName ?? "关联主表记录";
            // 添加关联主表字段（用于关联）
            var headerRow = new Row();
            uint columnIndex = 1;
            // 添加主表关联字段
            var relationCell = CreateTextCell(columnIndex++, 1, relationFieldName);
            headerRow.AppendChild(relationCell);
            foreach (var field in itemFields)
            {
                var cell = CreateTextCell(columnIndex++, 1, field.EffectiveColumnName);
                headerRow.AppendChild(cell);
            }
            sheetData.AppendChild(headerRow);
            // 添加明细数据行
            uint rowIndex = 2;
            int totalDetails = 0;
            foreach (var parentItem in data)
            {
                // 获取主表的关联值（订单编号）
                var parentKey = GetParentKeyValue(parentItem);
                // 获取集合数据
                var collection = collectionField.PropertyInfo.GetValue(parentItem);
                if (collection == null)
                {
                    continue;
                }
                // 尝试作为 XPCollection 处理
                var countProperty = collection.GetType().GetProperty("Count");
                if (countProperty == null)
                {
                    continue;
                }
                int count = (int)countProperty.GetValue(collection);
                // 使用枚举器遍历集合
                // XPCollection 实现了 IEnumerable 接口，我们可以直接使用 foreach
                System.Collections.IEnumerable enumerableCollection = collection as System.Collections.IEnumerable;
                if (enumerableCollection == null)
                {
                    continue;
                }
                foreach (var item in enumerableCollection)
                {
                    if (item == null) continue;
                    totalDetails++;
                    var dataRow = new Row() { RowIndex = rowIndex++ };
                    columnIndex = 1;
                    // 添加主表关联值
                    relationCell = CreateTextCell(columnIndex++, rowIndex - 1, parentKey);
                    dataRow.AppendChild(relationCell);
                    // 添加明细项字段
                    foreach (var field in itemFields)
                    {
                        try
                        {
                            var propValue = field.PropertyInfo.GetValue(item);
                            var convertResult = _dataConverter.ConvertToExcel(propValue, field);
                            string cellValue = convertResult?.ConvertedValue?.ToString() ?? "";
                            var cell = CreateTextCell(columnIndex++, rowIndex - 1, cellValue);
                            dataRow.AppendChild(cell);
                        }
                        catch
                        {
                            var cell = CreateTextCell(columnIndex++, rowIndex - 1, "");
                            dataRow.AppendChild(cell);
                        }
                    }
                    sheetData.AppendChild(dataRow);
                }
            }
        }
        /// <summary>
        /// 获取主表的关联键值（用于关联明细数据）
        /// </summary>
        private string GetParentKeyValue(object parentItem)
        {
            if (parentItem == null) return "";
            // 尝试获取具有 DefaultProperty 的属性值
            var parentType = parentItem.GetType();
            var defaultPropertyAttr = parentType.GetCustomAttributes(typeof(System.ComponentModel.DefaultPropertyAttribute), false)
                .FirstOrDefault() as System.ComponentModel.DefaultPropertyAttribute;
            if (defaultPropertyAttr != null && !string.IsNullOrEmpty(defaultPropertyAttr.Name))
            {
                var prop = parentType.GetProperty(defaultPropertyAttr.Name);
                if (prop != null)
                {
                    return prop.GetValue(parentItem)?.ToString() ?? parentItem.ToString();
                }
            }
            // 尝试获取名为 "Name"、"Title"、"Code" 或 "订单编号" 的属性
            var nameProp = parentType.GetProperty("Name") ??
                          parentType.GetProperty("Title") ??
                          parentType.GetProperty("Code") ??
                          parentType.GetProperty("订单编号");
            if (nameProp != null)
            {
                return nameProp.GetValue(parentItem)?.ToString() ?? parentItem.ToString();
            }
            return parentItem.ToString();
        }
        /// <summary>
        /// 创建文本单元格
        /// </summary>
        private Cell CreateTextCell(uint columnIndex, uint rowIndex, string text)
        {
            var cell = new Cell()
            {
                CellReference = GetCellReference(columnIndex, rowIndex),
                DataType = CellValues.String
            };
            cell.CellValue = new CellValue(text);
            return cell;
        }
        /// <summary>
        /// 获取单元格引用（如 A1, B2）
        /// </summary>
        private string GetCellReference(uint column, uint row)
        {
            string columnName = "";
            while (column > 0)
            {
                column--;
                columnName = Convert.ToChar('A' + (column % 26)) + columnName;
                column /= 26;
            }
            return columnName + row.ToString();
        }
        /// <summary>
        /// 解析文件内容
        /// </summary>
        private async Task<FileParseResult> ParseFileContentAsync(byte[] fileContent, Type objectType, ExcelImportOptions options, string fileName = null)
        {
            return await Task.Run(() =>
            {
                var result = new FileParseResult();
                try
                {
                    // 检测文件格式 - 使用实际文件名或默认名
                    var validation = ValidateFile(fileName ?? "temp.xlsx", fileContent);
                    if (!validation.IsValid)
                    {
                        result.ErrorMessage = validation.ErrorMessage;
                        return result;
                    }
                    // 根据格式解析内容
                    switch (validation.DetectedFormat)
                    {
                        case ExcelFormat.Csv:
                            result.DataRows = ParseCsvContent(fileContent, objectType, options);
                            break;
                        case ExcelFormat.Xlsx:
                        case ExcelFormat.Xls:
                            result.DataRows = ParseExcelContent(fileContent, objectType, options);
                            break;
                        default:
                            result.ErrorMessage = "不支持的文件格式";
                            return result;
                    }
                    result.IsSuccess = result.DataRows.Count > 0;
                    if (!result.IsSuccess && string.IsNullOrEmpty(result.ErrorMessage))
                    {
                        result.ErrorMessage = "文件中没有找到有效数据";
                    }
                    return result;
                }
                catch (Exception ex)
                {
                    result.ErrorMessage = $"文件解析失败: {ex.Message}";
                    return result;
                }
            });
        }
        /// <summary>
        /// 处理导入行
        /// </summary>
        private async Task<ImportRowResult> ProcessImportRowAsync(Dictionary<string, string> dataRow, Type objectType, IObjectSpace objectSpace, Dictionary<string, ExcelFieldConfiguration> importFields, ExcelImportOptions options, ExcelConfiguration config, int rowNumber)
        {
            return await Task.Run(() =>
            {
                var result = new ImportRowResult();
                try
                {
                    // 验证数据行
                    var validationResults = _dataValidator.ValidateRow(dataRow, importFields.Values.ToList(), rowNumber);
                    var validationErrors = validationResults.Where(v => !v.IsValid).ToList();
                    if (validationErrors.Any())
                    {
                        foreach (var error in validationErrors)
                        {
                            result.Errors.Add(error.ToExcelImportError());
                        }
                        // 如果是严格模式，直接返回错误
                        if (config.ClassConfiguration?.ValidationMode == Enums.ValidationMode.Strict)
                        {
                            return result;
                        }
                    }
                    // 根据导入模式处理
                    object targetObject = null;
                    switch (options.Mode)
                    {
                        case ExcelImportMode.CreateOnly:
                            // **修复**: 先检查是否存在,如果存在则跳过
                            targetObject = FindExistingObject(objectSpace, objectType, dataRow, importFields);
                            if (targetObject != null)
                            {
                                // 记录已存在,跳过该行
                                System.Diagnostics.Debug.WriteLine($"[CreateOnly] 行{rowNumber}: 记录已存在,跳过 - {GetKeyFieldsDisplay(dataRow, importFields)}");
                                result.Warnings.Add(new ExcelImportWarning
                                {
                                    RowNumber = rowNumber,
                                    FieldName = "数据跳过",
                                    WarningMessage = $"记录已存在(唯一标识: {GetKeyFieldsDisplay(dataRow, importFields)}),已跳过"
                                });
                                result.IsSuccess = true; // **标记为成功** - 跳过也算成功处理
                                result.CreatedObject = null; // **明确**: 没有创建对象
                                return result; // 返回成功(但有警告),跳过此行
                            }
                            // 不存在则创建新对象
                            System.Diagnostics.Debug.WriteLine($"[CreateOnly] 行{rowNumber}: 记录不存在,创建新对象 - {GetKeyFieldsDisplay(dataRow, importFields)}");
                            targetObject = objectSpace.CreateObject(objectType);
                            break;
                        case ExcelImportMode.UpdateOnly:
                            // 查找现有对象
                            System.Diagnostics.Debug.WriteLine($"[UpdateOnly] 行{rowNumber}: 开始UpdateOnly模式检查");
                            targetObject = FindExistingObject(objectSpace, objectType, dataRow, importFields);
                            if (targetObject == null)
                            {
                                // **DEBUG**: 记录不存在,跳过该行(不创建新记录)
                                System.Diagnostics.Debug.WriteLine($"[UpdateOnly] 行{rowNumber}: 记录不存在,跳过 - {GetKeyFieldsDisplay(dataRow, importFields)}");

                                result.Warnings.Add(new ExcelImportWarning
                                {
                                    RowNumber = rowNumber,
                                    FieldName = "数据跳过",
                                    WarningMessage = $"数据库中不存在该记录(唯一标识: {GetKeyFieldsDisplay(dataRow, importFields)}),已跳过"
                                });
                                result.IsSuccess = true; // **标记为成功** - 跳过也算成功处理
                                result.CreatedObject = null; // **明确**: 没有创建对象
                                System.Diagnostics.Debug.WriteLine($"[UpdateOnly] 行{rowNumber}: 准备返回null,警告数={result.Warnings.Count}");
                                return result; // 返回成功(但有警告),跳过此行
                            }
                            // **DEBUG**: 找到记录,准备更新
                            System.Diagnostics.Debug.WriteLine($"[UpdateOnly] 行{rowNumber}: 找到记录,准备更新 - {GetKeyFieldsDisplay(dataRow, importFields)}");
                            break;
                        case ExcelImportMode.CreateOrUpdate:
                            targetObject = FindExistingObject(objectSpace, objectType, dataRow, importFields);
                            if (targetObject == null)
                            {
                                System.Diagnostics.Debug.WriteLine($"[CreateOrUpdate] 行{rowNumber}: 记录不存在,创建新记录 - {GetKeyFieldsDisplay(dataRow, importFields)}");
                                targetObject = objectSpace.CreateObject(objectType);
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"[CreateOrUpdate] 行{rowNumber}: 记录已存在,更新记录 - {GetKeyFieldsDisplay(dataRow, importFields)}");
                            }
                            break;
                        case ExcelImportMode.ReplaceAll:
                            // 已经在导入开始前删除了所有记录,所以这里直接创建新对象
                            targetObject = objectSpace.CreateObject(objectType);
                            break;
                        default:
                            targetObject = objectSpace.CreateObject(objectType);
                            break;
                    }
                    // 设置属性值
                    foreach (var kvp in dataRow)
                    {
                        if (importFields.TryGetValue(kvp.Key, out var fieldConfig))
                        {
                            try
                            {
                                // 直接使用 fieldConfig.PropertyInfo，避免类型不匹配问题
                                var property = fieldConfig.PropertyInfo;
                                if (property != null && property.CanWrite)
                                {
                                    var convertResult = _dataConverter.ConvertFromExcel(kvp.Value, fieldConfig);
                                    if (convertResult.IsSuccess)
                                    {
                                        // 处理空值情况：如果字段不是必需的，且转换后的值为空或默认值，则提供默认值以通过验证
                                        object valueToSet = convertResult.ConvertedValue;
                                        // DataDictionaryItem 类型特殊处理 - 需要查询数据库获取对象
                                        if (property.PropertyType.Name == "DataDictionaryItem" || property.PropertyType.FullName?.Contains("DataDictionaryItem") == true)
                                        {
                                            if (valueToSet == null || string.IsNullOrWhiteSpace(valueToSet.ToString()))
                                            {
                                                // 空值，跳过设置
                                                continue;
                                            }
                                            // 查询 DataDictionaryItem 对象
                                            try
                                            {
                                                var dictionaryName = GetDictionaryNameFromProperty(property);
                                                var itemValue = valueToSet.ToString();
                                                var dataDictItem = FindDataDictionaryItem(objectSpace, dictionaryName, itemValue);
                                                if (dataDictItem != null)
                                                {
                                                    valueToSet = dataDictItem;
                                                    var itemName = dataDictItem.GetType().GetProperty("Name")?.GetValue(dataDictItem)?.ToString() ?? dataDictItem.ToString();
                                                }
                                                else
                                                {
                                                    // 未找到字典项,自动创建
                                                    dataDictItem = CreateDataDictionaryItem(objectSpace, dictionaryName, itemValue);
                                                    if (dataDictItem != null)
                                                    {
                                                        valueToSet = dataDictItem;
                                                        var itemName = dataDictItem.GetType().GetProperty("Name")?.GetValue(dataDictItem)?.ToString() ?? dataDictItem.ToString();
                                                        // 添加创建成功的信息
                                                        result.Warnings.Add(new ExcelImportWarning
                                                        {
                                                            RowNumber = rowNumber,
                                                            FieldName = kvp.Key,
                                                            WarningMessage = $"自动创建字典项: {dictionaryName} - {itemValue}",
                                                            OriginalValue = itemValue,
                                                            ConvertedValue = itemName
                                                        });
                                                    }
                                                    else
                                                    {
                                                        // 创建失败,添加警告
                                                        result.Warnings.Add(new ExcelImportWarning
                                                        {
                                                            RowNumber = rowNumber,
                                                            FieldName = kvp.Key,
                                                            WarningMessage = $"字典项不存在且创建失败: {dictionaryName} - {itemValue}",
                                                            OriginalValue = itemValue
                                                        });
                                                        // 跳过设置此字段
                                                        continue;
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                result.Errors.Add(new ExcelImportError
                                                {
                                                    RowNumber = rowNumber,
                                                    FieldName = kvp.Key,
                                                    ErrorMessage = $"查找字典项失败: {ex.Message}",
                                                    OriginalValue = kvp.Value,
                                                    ErrorType = ExcelImportErrorType.SystemError
                                                });
                                                continue;
                                            }
                                        }
                                        else if (!fieldConfig.FieldAttribute.IsRequired)
                                        {
                                            // 对于字符串类型，空字符串可以改为null
                                            if (property.PropertyType == typeof(string) && valueToSet is string strValue && string.IsNullOrEmpty(strValue))
                                            {
                                                valueToSet = null;
                                            }
                                            // 对于数值类型，如果是0且不在Excel中提供，可以跳过设置
                                            else if ((property.PropertyType == typeof(long) || property.PropertyType == typeof(int)) &&
                                                    (valueToSet is long || valueToSet is int) &&
                                                    Convert.ToInt64(valueToSet) == 0 &&
                                                    string.IsNullOrWhiteSpace(kvp.Value))
                                            {
                                                // 跳过设置此字段，使用数据库默认值
                                                continue;
                                            }
                                            // **修复**: 对于 DateTime 类型，如果是 null 且 Excel 值为空，跳过设置
                                            else if (property.PropertyType == typeof(DateTime) &&
                                                    valueToSet == null &&
                                                    string.IsNullOrWhiteSpace(kvp.Value))
                                            {
                                                // 跳过设置此字段，使用数据库默认值
                                                continue;
                                            }
                                        }
                                        // **修复**: 如果 valueToSet 为 DateTime.MinValue 且原始 Excel 值为空，也跳过设置
                                        if (property.PropertyType == typeof(DateTime) &&
                                            valueToSet is DateTime dateValue &&
                                            dateValue == DateTime.MinValue &&
                                            string.IsNullOrWhiteSpace(kvp.Value))
                                        {
                                            continue;
                                        }
                                        // 直接使用原始的 PropertyInfo
                                        property.SetValue(targetObject, valueToSet);
                                        if (convertResult.HasWarning)
                                        {
                                            result.Warnings.Add(convertResult.ToExcelImportWarning(rowNumber, kvp.Key));
                                        }
                                    }
                                    else
                                    {
                                        result.Errors.Add(convertResult.ToExcelImportError(rowNumber, kvp.Key));
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                result.Errors.Add(new ExcelImportError
                                {
                                    RowNumber = rowNumber,
                                    FieldName = kvp.Key,
                                    ErrorMessage = ex.Message,
                                    OriginalValue = kvp.Value,
                                    ErrorType = ExcelImportErrorType.SystemError
                                });
                            }
                        }
                    }
                    result.IsSuccess = result.Errors.Count == 0;
                    result.CreatedObject = targetObject; // 保存创建的对象
                    return result;
                }
                catch (Exception ex)
                {
                    result.Errors.Add(new ExcelImportError
                    {
                        RowNumber = rowNumber,
                        FieldName = "行处理",
                        ErrorMessage = ex.Message,
                        ErrorType = ExcelImportErrorType.SystemError
                    });
                    return result;
                }
            });
        }
        /// <summary>
        /// 查找现有对象
        /// </summary>
        private object FindExistingObject(IObjectSpace objectSpace, Type objectType, Dictionary<string, string> dataRow, Dictionary<string, ExcelFieldConfiguration> importFields)
        {
            try
            {
                // **优先查找带有 RuleUniqueValue 特性的字段**
                string keyPropertyName = null;
                foreach (var prop in objectType.GetProperties())
                {
                    // 检查属性是否有 RuleUniqueValue 特性
                    var hasUniqueRule = prop.GetCustomAttributes(false)
                        .Any(attr => attr.GetType().Name.Contains("RuleUniqueValue"));
                    if (hasUniqueRule)
                    {
                        keyPropertyName = prop.Name;
                        System.Diagnostics.Debug.WriteLine($"[FindExistingObject] 找到唯一字段: {keyPropertyName}");
                        break;
                    }
                }

                // 如果没有找到唯一字段,使用 DefaultProperty
                if (string.IsNullOrEmpty(keyPropertyName))
                {
                    var defaultPropertyAttr = objectType.GetCustomAttributes(typeof(System.ComponentModel.DefaultPropertyAttribute), false)
                        .FirstOrDefault() as System.ComponentModel.DefaultPropertyAttribute;
                    if (defaultPropertyAttr != null && !string.IsNullOrEmpty(defaultPropertyAttr.Name))
                    {
                        keyPropertyName = defaultPropertyAttr.Name;
                        System.Diagnostics.Debug.WriteLine($"[FindExistingObject] 使用DefaultProperty: {keyPropertyName}");
                    }
                }

                // 如果还是没有,尝试常见的属性名
                if (string.IsNullOrEmpty(keyPropertyName))
                {
                    var commonKeyNames = new[] { "Name", "Title", "Code", "订单编号", "OrderNo", "Oid", "员工编号" };
                    foreach (var name in commonKeyNames)
                    {
                        if (objectType.GetProperty(name) != null)
                        {
                            keyPropertyName = name;
                            System.Diagnostics.Debug.WriteLine($"[FindExistingObject] 使用常见字段: {keyPropertyName}");
                            break;
                        }
                    }
                }

                if (string.IsNullOrEmpty(keyPropertyName))
                {
                    System.Diagnostics.Debug.WriteLine($"[FindExistingObject] 未找到唯一字段,返回null");
                    return null;
                }
                // 在配置中查找对应的Excel列名
                var keyFieldConfig = importFields.Values.FirstOrDefault(f => f.PropertyInfo.Name == keyPropertyName);
                string keyColumnName = keyFieldConfig?.EffectiveColumnName ?? keyPropertyName;
                System.Diagnostics.Debug.WriteLine($"[FindExistingObject] Excel列名: {keyColumnName}, 属性名: {keyPropertyName}");

                // 从数据行中获取键值
                if (!dataRow.TryGetValue(keyColumnName, out var keyValue) || string.IsNullOrEmpty(keyValue))
                {
                    System.Diagnostics.Debug.WriteLine($"[FindExistingObject] Excel中未找到键值: {keyColumnName}");
                    return null;
                }

                System.Diagnostics.Debug.WriteLine($"[FindExistingObject] 查找键值: {keyPropertyName}={keyValue}");

                // 使用ObjectSpace查询现有对象
                var criteria = CriteriaOperator.Parse($"[{keyPropertyName}] = ?", keyValue);
                var existingObjects = objectSpace.GetObjects(objectType, criteria);
                var existingObject = existingObjects.Cast<object>().FirstOrDefault();

                System.Diagnostics.Debug.WriteLine($"[FindExistingObject] 查询结果: {(existingObject != null ? "找到记录" : "未找到记录")}");

                return existingObject;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[FindExistingObject] 查找失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 获取关键字段的显示值(用于警告消息)
        /// </summary>
        private string GetKeyFieldsDisplay(Dictionary<string, string> dataRow, Dictionary<string, ExcelFieldConfiguration> importFields)
        {
            try
            {
                var keyValues = new List<string>();

                // **第一优先级**: 只获取带有 RuleUniqueValue 的字段(最多2个)
                foreach (var fieldConfig in importFields.Values.OrderBy(f => f.EffectiveSortOrder))
                {
                    if (keyValues.Count >= 2) break; // 最多显示2个唯一字段

                    var columnName = fieldConfig.EffectiveColumnName;
                    if (dataRow.ContainsKey(columnName) && !string.IsNullOrEmpty(dataRow[columnName]))
                    {
                        // 检查是否是唯一字段
                        var hasUniqueRule = fieldConfig.PropertyInfo?.GetCustomAttributes(false)
                            .Any(attr => attr.GetType().Name.Contains("RuleUniqueValue")) == true;

                        if (hasUniqueRule)
                        {
                            keyValues.Add($"{columnName}={dataRow[columnName]}");
                        }
                    }
                }

                // **如果没有 RuleUniqueValue 字段,使用 DefaultProperty**
                if (keyValues.Count == 0)
                {
                    var defaultFieldConfig = importFields.Values
                        .OrderBy(f => f.EffectiveSortOrder)
                        .FirstOrDefault(f => !string.IsNullOrEmpty(dataRow[f.EffectiveColumnName]));

                    if (defaultFieldConfig != null)
                    {
                        var columnName = defaultFieldConfig.EffectiveColumnName;
                        keyValues.Add($"{columnName}={dataRow[columnName]}");
                    }
                }

                return keyValues.Count > 0 ? string.Join(", ", keyValues) : "未知";
            }
            catch
            {
                return "未知";
            }
        }

        /// <summary>
        /// 从属性获取 DataDictionary 名称
        /// </summary>
        private string GetDictionaryNameFromProperty(PropertyInfo property)
        {
            try
            {
                // 查找 DataDictionary 特性
                var dataDictAttr = property.GetCustomAttributes(true)
                    .FirstOrDefault(a =>
                    {
                        var attrType = a.GetType();
                        return attrType.Name == "DataDictionaryAttribute" ||
                               attrType.FullName?.Contains("DataDictionaryAttribute") == true;
                    });

                if (dataDictAttr != null)
                {
                    // 获取 DataDictionaryName 属性 (修正属性名)
                    var dictNameProp = dataDictAttr.GetType().GetProperty("DataDictionaryName");
                    if (dictNameProp != null)
                    {
                        var dictName = dictNameProp.GetValue(dataDictAttr)?.ToString();
                        if (!string.IsNullOrEmpty(dictName))
                        {
                            return dictName;
                        }
                    }

                    // 如果找不到 DataDictionaryName，尝试通过反射获取所有属性
                    var props = dataDictAttr.GetType().GetProperties();
                    foreach (var prop in props)
                    {
                        if (prop.Name.Contains("DictionaryName") || prop.Name.Contains("Name"))
                        {
                            var value = prop.GetValue(dataDictAttr)?.ToString();
                            if (!string.IsNullOrEmpty(value) && value != property.Name)
                            {
                                return value;
                            }
                        }
                    }
                }
                // 如果没有找到特性或无法获取字典名称，记录警告并使用属性名
                System.Diagnostics.Debug.WriteLine($"[GetDictionaryNameFromProperty] 警告: 属性 {property.Name} 未找到 DataDictionary 特性或 DataDictionaryName 值");
                return property.Name;
            }
            catch
            {
                return property.Name;
            }
        }
        /// <summary>
        /// 查找 DataDictionaryItem 对象
        /// </summary>
        private object FindDataDictionaryItem(IObjectSpace objectSpace, string dictionaryName, string itemValue)
        {
            try
            {
                // **优先检查缓存** (本次导入中已创建的项)
                var cacheKey = $"{dictionaryName}:{itemValue}";
                if (_dataDictionaryItemCache.ContainsKey(cacheKey))
                {
                    return _dataDictionaryItemCache[cacheKey];
                }

                // 方法1: 尝试查找 DataDictionaryItem 类型
                var dataDictItemType = Type.GetType("Wxy.Xaf.DataDictionary.DataDictionaryItem, Wxy.Xaf.DataDictionary");
                if (dataDictItemType == null)
                {
                    // 尝试在加载的程序集中查找
                    foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
                    {
                        dataDictItemType = assembly.GetType("Wxy.Xaf.DataDictionary.DataDictionaryItem");
                        if (dataDictItemType != null)
                        {
                            break;
                        }
                    }
                }
                if (dataDictItemType == null)
                {
                    return null;
                }
                // 方法2: 尝试通过 Name 属性和 DataDictionary.Name 查找
                var criteriaByName = CriteriaOperator.Parse($"[Name] = ? AND [DataDictionary.Name] = ?", itemValue, dictionaryName);
                var objectsByName = objectSpace.GetObjects(dataDictItemType, criteriaByName);
                var result = objectsByName.Cast<object>().FirstOrDefault();
                if (result != null)
                {
                    // 加入缓存
                    _dataDictionaryItemCache[cacheKey] = result;
                    return result;
                }
                // 方法3: 尝试通过 Code 属性查找
                var criteriaByCode = CriteriaOperator.Parse($"[Code] = ? AND [DataDictionary.Name] = ?", itemValue, dictionaryName);
                var objectsByCode = objectSpace.GetObjects(dataDictItemType, criteriaByCode);
                result = objectsByCode.Cast<object>().FirstOrDefault();
                if (result != null)
                {
                    // 加入缓存
                    _dataDictionaryItemCache[cacheKey] = result;
                    return result;
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
        /// <summary>
        /// 创建 DataDictionaryItem 对象
        /// </summary>
        private object CreateDataDictionaryItem(IObjectSpace objectSpace, string dictionaryName, string itemValue)
        {
            try
            {
                // **首先检查缓存**
                var cacheKey = $"{dictionaryName}:{itemValue}";
                if (_dataDictionaryItemCache.ContainsKey(cacheKey))
                {
                    return _dataDictionaryItemCache[cacheKey];
                }

                // **查找或创建 DataDictionary**
                object dataDictionary;
                if (_dataDictionaryCache.ContainsKey(dictionaryName))
                {
                    dataDictionary = _dataDictionaryCache[dictionaryName];
                }
                else
                {
                    // 查找 DataDictionary 类型
                    var dataDictType = Type.GetType("Wxy.Xaf.DataDictionary.DataDictionary, Wxy.Xaf.DataDictionary");
                    if (dataDictType == null)
                    {
                        foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
                        {
                            dataDictType = assembly.GetType("Wxy.Xaf.DataDictionary.DataDictionary");
                            if (dataDictType != null)
                            {
                                break;
                            }
                        }
                    }
                    if (dataDictType == null)
                    {
                        return null;
                    }

                    // 查找 DataDictionary
                    var criteria = CriteriaOperator.Parse($"[Name] = ?", dictionaryName);
                    var dataDicts = objectSpace.GetObjects(dataDictType, criteria);
                    dataDictionary = dataDicts.Cast<object>().FirstOrDefault();

                    if (dataDictionary == null)
                    {
                        // **DataDictionary 不存在,尝试创建**
                        System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] 数据字典 '{dictionaryName}' 不存在,尝试创建");

                        // 尝试使用 XPO 直接创建 DataDictionary
                        if (objectSpace.GetType().FullName == "DevExpress.ExpressApp.Xpo.XPObjectSpace")
                        {
                            try
                            {
                                var sessionProperty = objectSpace.GetType().GetProperty("Session");
                                var session = sessionProperty?.GetValue(objectSpace);

                                if (session != null)
                                {
                                    var ctor = dataDictType.GetConstructor(new[] { session.GetType() });
                                    dataDictionary = ctor?.Invoke(new[] { session });

                                    if (dataDictionary != null)
                                    {
                                        var nameProp = dataDictType.GetProperty("Name");
                                        nameProp?.SetValue(dataDictionary, dictionaryName);

                                        var saveMethod = dataDictionary.GetType().GetMethod("Save");
                                        saveMethod?.Invoke(dataDictionary, null);

                                        System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] ✓ 成功创建数据字典(XPO): {dictionaryName}");
                                    }
                                }
                            }
                            catch (Exception xpoEx)
                            {
                                System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] 创建数据字典失败(XPO): {xpoEx.Message}");
                                System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] 堆栈: {xpoEx.StackTrace}");
                            }
                        }

                        // **不再使用 ObjectSpace 方式**,因为在 Blazor 中会导致跨线程问题
                        // 如果 XPO 创建失败,直接返回 null
                        if (dataDictionary == null)
                        {
                            System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] 无法创建数据字典 '{dictionaryName}' (XPO创建失败)");
                            return null;
                        }
                    }

                    // **加入 DataDictionary 缓存**
                    _dataDictionaryCache[dictionaryName] = dataDictionary;
                }

                // 创建 DataDictionaryItem
                var dataDictItemType = Type.GetType("Wxy.Xaf.DataDictionary.DataDictionaryItem, Wxy.Xaf.DataDictionary");
                if (dataDictItemType == null)
                {
                    foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
                    {
                        dataDictItemType = assembly.GetType("Wxy.Xaf.DataDictionary.DataDictionaryItem");
                        if (dataDictItemType != null)
                        {
                            break;
                        }
                    }
                }
                if (dataDictItemType == null)
                {
                    return null;
                }

                // **优先检查缓存**: 首先检查本次导入中已经创建过的项
                if (_dataDictionaryItemCache.TryGetValue(cacheKey, out var cachedItem))
                {
                    System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] 从缓存中找到项: {dictionaryName} - {itemValue}");
                    return cachedItem;
                }

                // **重要**: 在创建之前先检查是否已存在同名项 (数据库中已存在的)
                // 先使用 ObjectSpace 查询
                var existingItemCriteria = CriteriaOperator.Parse($"[Name] = ? AND [DataDictionary.Name] = ?", itemValue, dictionaryName);
                var existingItems = objectSpace.GetObjects(dataDictItemType, existingItemCriteria);
                var existingItem = existingItems.Cast<object>().FirstOrDefault();
                if (existingItem != null)
                {
                    // 加入缓存并返回
                    _dataDictionaryItemCache[cacheKey] = existingItem;
                    System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] ObjectSpace找到已存在项: {dictionaryName} - {itemValue}");
                    return existingItem;
                }

                // 确实不存在才创建新项

                // **检查 objectSpace 类型,尝试使用 XPObjectSpace 直接操作**
                var objectSpaceTypeName = objectSpace.GetType().FullName;
                System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] ObjectSpace类型: {objectSpaceTypeName}");

                // 支持标准 XPObjectSpace 和 SecuredXPObjectSpace (安全包装版本)
                if (objectSpaceTypeName == "DevExpress.ExpressApp.Xpo.XPObjectSpace" ||
                    objectSpaceTypeName == "DevExpress.ExpressApp.Security.SecuredXPObjectSpace")
                {
                    try
                    {
                        // 使用反射获取 Session 属性,避免硬引用
                        var sessionProperty = objectSpace.GetType().GetProperty("Session");
                        var session = sessionProperty?.GetValue(objectSpace);

                        if (session != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] 获取到 Session: {session.GetType().FullName}");

                            // 使用反射获取构造函数: DataDictionaryItem(Session)
                            var ctor = dataDictItemType.GetConstructor(new[] { session.GetType() });
                            if (ctor == null)
                            {
                                System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] 无法找到构造函数 {dataDictItemType.Name}(Session)");
                            }
                            else
                            {
                                var xpNewItem = ctor?.Invoke(new[] { session });

                                if (xpNewItem != null)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] 成功创建字典项对象: {xpNewItem.GetType().FullName}");

                                    // **关键修复**: 先设置 Name 和 Code,暂不设置 DataDictionary 关联
                                    // 这可以避免 IsNameUnique 验证规则在 Save 时检查 DataDictionary.Items 集合
                                    var itemNameProp = dataDictItemType.GetProperty("Name");
                                    itemNameProp?.SetValue(xpNewItem, itemValue);
                                    var codeProp = dataDictItemType.GetProperty("Code");
                                    codeProp?.SetValue(xpNewItem, itemValue);

                                    System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] 基本属性设置完成，准备保存(不设置关联): {dictionaryName} - {itemValue}");

                                    // 先保存对象(此时 DataDictionary 为 null,所以 IsNameUnique 验证会通过)
                                    try
                                    {
                                        var saveMethod = xpNewItem.GetType().GetMethod("Save", System.Type.EmptyTypes);
                                        saveMethod?.Invoke(xpNewItem, null);
                                        System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] ✓ 成功保存字典项(XPO,无关联): {dictionaryName} - {itemValue}");
                                    }
                                    catch (Exception saveEx)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] ✗ 第一次保存失败: {dictionaryName} - {itemValue}");
                                        System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem]   错误: {saveEx.Message}");
                                        throw;
                                    }

                                    // **保存后再设置 DataDictionary 关联** (此时不会触发 Save,所以不会触发验证)
                                    var dataDictProp = dataDictItemType.GetProperty("DataDictionary");
                                    dataDictProp?.SetValue(xpNewItem, dataDictionary);
                                    System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] ✓ 已设置 DataDictionary 关联: {dictionaryName} - {itemValue}");

                                    // **再次保存以提交关联关系**
                                    try
                                    {
                                        var saveMethod = xpNewItem.GetType().GetMethod("Save", System.Type.EmptyTypes);
                                        saveMethod?.Invoke(xpNewItem, null);
                                        System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] ✓ 成功更新关联关系(XPO): {dictionaryName} - {itemValue}");
                                    }
                                    catch (Exception saveEx2)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] ✗ 第二次保存失败: {dictionaryName} - {itemValue}");
                                        System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem]   错误: {saveEx2.Message}");
                                        throw;
                                    }

                                    // **加入缓存**
                                    _dataDictionaryItemCache[cacheKey] = xpNewItem;
                                    return xpNewItem;
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] 构造函数返回null");
                                }
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] 无法获取 Session");
                        }
                    }
                    catch (Exception xpoEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] XPO创建字典项失败: {dictionaryName} - {itemValue}");
                        System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] XPO错误: {xpoEx.Message}");
                        System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] XPO堆栈: {xpoEx.StackTrace}");
                        // **不再尝试 ObjectSpace 方式**,因为在 Blazor 中会导致跨线程问题
                        return null;
                    }
                }

                // **不再使用 ObjectSpace 方式创建字典项**,因为在 Blazor 中会导致跨线程问题
                // 直接返回 null,表示创建失败
                System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] ObjectSpace类型不匹配({objectSpaceTypeName}),无法创建字典项: {dictionaryName} - {itemValue}");
                System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] 期望类型: DevExpress.ExpressApp.Xpo.XPObjectSpace");
                return null;
            }
            catch (Exception ex)
            {
                // 记录详细的错误信息以便调试
                System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] 创建字典项失败: {dictionaryName} - {itemValue}");
                System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] 错误详情: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[CreateDataDictionaryItem] 堆栈跟踪: {ex.StackTrace}");
                return null;
            }
        }
        /// <summary>
        /// 查找现有明细对象（考虑父对象关联）
        /// </summary>
        private object FindExistingDetailObject(IObjectSpace objectSpace, Type detailType, Dictionary<string, string> dataRow, Dictionary<string, ExcelFieldConfiguration> detailFieldDict, object parentObject)
        {
            try
            {
                // 获取明细对象的默认属性（主键字段）
                var defaultPropertyAttr = detailType.GetCustomAttributes(typeof(System.ComponentModel.DefaultPropertyAttribute), false)
                    .FirstOrDefault() as System.ComponentModel.DefaultPropertyAttribute;
                string keyPropertyName = null;
                if (defaultPropertyAttr != null && !string.IsNullOrEmpty(defaultPropertyAttr.Name))
                {
                    keyPropertyName = defaultPropertyAttr.Name;
                }
                else
                {
                    // 尝试常见的属性名
                    var commonKeyNames = new[] { "Name", "Title", "Code", "订单编号", "OrderNo", "Oid" };
                    foreach (var name in commonKeyNames)
                    {
                        if (detailType.GetProperty(name) != null)
                        {
                            keyPropertyName = name;
                            break;
                        }
                    }
                }
                if (string.IsNullOrEmpty(keyPropertyName))
                {
                    return null;
                }
                // 在配置中查找对应的Excel列名
                var keyFieldConfig = detailFieldDict.Values.FirstOrDefault(f => f.PropertyInfo.Name == keyPropertyName);
                string keyColumnName = keyFieldConfig?.EffectiveColumnName ?? keyPropertyName;
                // 从数据行中获取键值
                if (!dataRow.TryGetValue(keyColumnName, out var keyValue) || string.IsNullOrEmpty(keyValue))
                {
                    return null;
                }
                // 构建查询条件，同时匹配键值和父对象
                // 首先需要找到关联到父对象的属性名
                string parentPropertyName = null;
                foreach (var prop in detailType.GetProperties())
                {
                    // 查找类型为父对象类型的属性
                    if (prop.PropertyType == parentObject.GetType())
                    {
                        parentPropertyName = prop.Name;
                        break;
                    }
                }
                if (string.IsNullOrEmpty(parentPropertyName))
                {
                    return null;
                }
                // 使用ObjectSpace查询现有明细对象，同时匹配键值和父对象
                var criteria = CriteriaOperator.Parse($"[{keyPropertyName}] = ? AND [{parentPropertyName}] = ?", keyValue, parentObject);
                var existingObjects = objectSpace.GetObjects(detailType, criteria);
                var existingObject = existingObjects.Cast<object>().FirstOrDefault();
                if (existingObject != null)
                {
                }
                else
                {
                }
                return existingObject;
            }
            catch
            {
                return null;
            }
        }
        /// <summary>
        /// 转换值类型（已废弃，使用DataConverter替代）
        /// </summary>
        [Obsolete("使用DataConverter.ConvertFromExcel替代")]
        private ValueConvertResult ConvertValue(string value, Type targetType, ExcelFieldConfiguration fieldConfig)
        {
            var result = new ValueConvertResult();
            try
            {
                if (string.IsNullOrEmpty(value))
                {
                    // 空值处理
                    if (targetType.IsGenericType && targetType.GetGenericTypeDefinition() == typeof(Nullable<>))
                    {
                        result.Value = null;
                        result.IsSuccess = true;
                        return result;
                    }
                    
                    if (targetType == typeof(string))
                    {
                        result.Value = string.Empty;
                        result.IsSuccess = true;
                        return result;
                    }
                    
                    if (targetType.IsValueType)
                    {
                        result.Value = Activator.CreateInstance(targetType);
                        result.IsSuccess = true;
                        result.HasWarning = true;
                        result.WarningMessage = "空值已转换为默认值";
                        return result;
                    }
                    
                    result.Value = null;
                    result.IsSuccess = true;
                    return result;
                }
                // 处理可空类型
                Type actualType = targetType;
                if (targetType.IsGenericType && targetType.GetGenericTypeDefinition() == typeof(Nullable<>))
                {
                    actualType = Nullable.GetUnderlyingType(targetType);
                }
                // 特殊类型处理
                if (actualType == typeof(DateTime))
                {
                    if (DateTime.TryParse(value, out DateTime dateValue))
                    {
                        result.Value = dateValue;
                        result.IsSuccess = true;
                        return result;
                    }
                    else
                    {
                        result.ErrorMessage = $"无法将 '{value}' 转换为日期时间格式";
                        return result;
                    }
                }
                if (actualType == typeof(bool))
                {
                    var lowerValue = value.ToLower().Trim();
                    if (lowerValue == "true" || lowerValue == "1" || lowerValue == "是" || lowerValue == "yes")
                    {
                        result.Value = true;
                        result.IsSuccess = true;
                        return result;
                    }
                    if (lowerValue == "false" || lowerValue == "0" || lowerValue == "否" || lowerValue == "no")
                    {
                        result.Value = false;
                        result.IsSuccess = true;
                        return result;
                    }
                    
                    result.ErrorMessage = $"无法将 '{value}' 转换为布尔值";
                    return result;
                }
                if (actualType.IsEnum)
                {
                    try
                    {
                        // 对值进行 trim，去除前后空格
                        var trimmedValue = value.Trim();
                        // 检查是否有枚举映射格式
                        if (!string.IsNullOrEmpty(fieldConfig.FieldAttribute.Format) &&
                            fieldConfig.FieldAttribute.Format.Contains("="))
                        {
                            var mappings = ParseEnumMappings(fieldConfig.FieldAttribute.Format);
                            // 创建反向映射：显示值 -> 枚举名
                            var reverseMappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                            foreach (var mapping in mappings)
                            {
                                reverseMappings[mapping.Value] = mapping.Key;
                            }
                            // 尝试通过显示值查找枚举名
                            if (reverseMappings.TryGetValue(trimmedValue, out string enumName))
                            {
                                if (Enum.IsDefined(actualType, enumName))
                                {
                                    result.Value = Enum.Parse(actualType, enumName, true);
                                    result.IsSuccess = true;
                                    return result;
                                }
                            }
                        }
                        // 直接解析枚举
                        result.Value = Enum.Parse(actualType, trimmedValue, true);
                        result.IsSuccess = true;
                        return result;
                    }
                    catch
                    {
                        result.ErrorMessage = $"无法将 '{value}' 转换为枚举类型 {actualType.Name}";
                        return result;
                    }
                }
                // 通用转换
                result.Value = Convert.ChangeType(value, actualType);
                result.IsSuccess = true;
                return result;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"类型转换失败: {ex.Message}";
                return result;
            }
        }
        /// <summary>
        /// 格式化单元格值
        /// </summary>
        private string FormatCellValue(object value)
        {
            if (value == null)
                return "";
            if (value is DateTime dateTime)
                return dateTime.ToString("yyyy-MM-dd HH:mm:ss");
            if (value is decimal || value is double || value is float)
                return value.ToString();
            return value.ToString();
        }
        /// <summary>
        /// 验证CSV内容
        /// </summary>
        private bool ValidateCsvContent(byte[] content)
        {
            try
            {
                // 使用编码检测来验证CSV内容
                var encodingResult = EncodingDetector.DetectEncoding(content);
                
                // 尝试解码内容
                string text;
                if (encodingResult.HasBom)
                {
                    var bomLength = encodingResult.DetectedEncoding == Encoding.UTF8 ? 3 : 
                                   encodingResult.DetectedEncoding == Encoding.Unicode || encodingResult.DetectedEncoding == Encoding.BigEndianUnicode ? 2 : 4;
                    text = encodingResult.DetectedEncoding.GetString(content, bomLength, content.Length - bomLength);
                }
                else
                {
                    text = encodingResult.DetectedEncoding.GetString(content);
                }
                
                // 检查是否为有效的CSV格式（包含逗号分隔符或至少有可读文本）
                return !string.IsNullOrWhiteSpace(text) && (text.Contains(",") || text.Length > 0);
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 验证XLSX内容
        /// </summary>
        private bool ValidateXlsxContent(byte[] content)
        {
            try
            {
                // 检查XLSX文件头
                return content.Length > 4 && 
                       content[0] == 0x50 && content[1] == 0x4B && 
                       content[2] == 0x03 && content[3] == 0x04;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 验证XLS内容
        /// </summary>
        private bool ValidateXlsContent(byte[] content)
        {
            try
            {
                // 检查XLS文件头
                return content.Length > 8 && 
                       content[0] == 0xD0 && content[1] == 0xCF && 
                       content[2] == 0x11 && content[3] == 0xE0;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 解析CSV内容
        /// </summary>
        private List<Dictionary<string, string>> ParseCsvContent(byte[] fileContent, Type objectType, ExcelImportOptions options)
        {
            var result = new List<Dictionary<string, string>>();
            try
            {
                // 使用增强的编码检测（保持兼容性，但推荐使用UTF-8）
                var encodingResult = EncodingDetector.DetectEncoding(fileContent);
                
                // 记录编码检测结果（用于调试）
                if (encodingResult.DetectedEncoding != Encoding.UTF8)
                {
                }
                if (encodingResult.TriedEncodings.Count > 0)
                {
                }
                // 根据检测结果解码内容
                string content = DecodeContentWithFallback(fileContent, encodingResult);
                // 验证解码结果
                if (string.IsNullOrEmpty(content))
                {
                    throw new Exception("文件内容解码后为空，可能是编码检测错误");
                }
                // 检查并记录中文字符信息
                bool hasChineseContent = EncodingDetector.ContainsChineseCharacters(content);
                if (hasChineseContent)
                {
                    var chineseRatio = EncodingDetector.CalculateChineseCharacterRatio(content);
                    
                    // 输出前100个字符用于调试
                    var preview = content.Length > 100 ? content.Substring(0, 100) : content;
                }
                else
                {
                    // 输出前100个字符用于调试
                    var preview = content.Length > 100 ? content.Substring(0, 100) : content;
                    
                    // 如果没有检测到中文字符但内容包含乱码符号，尝试其他编码
                    if (content.Contains("�"))
                    {
                        content = TryAlternativeEncodings(fileContent);
                        
                        // 重新检查中文字符
                        hasChineseContent = EncodingDetector.ContainsChineseCharacters(content);
                        if (hasChineseContent)
                        {
                            var chineseRatio = EncodingDetector.CalculateChineseCharacterRatio(content);
                        }
                        
                        var newPreview = content.Length > 100 ? content.Substring(0, 100) : content;
                    }
                }
                var lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length < (options.HasHeaderRow ? 2 : 1))
                {
                    return result;
                }
                List<string> headers;
                int dataStartIndex;
                if (options.HasHeaderRow)
                {
                    headers = ParseCsvLine(lines[0]);
                    dataStartIndex = 1;
                }
                else
                {
                    // 如果没有表头，使用第一行数据生成默认表头
                    var firstRowData = ParseCsvLine(lines[0]);
                    headers = firstRowData.Select((_, index) => $"Column{index + 1}").ToList();
                    dataStartIndex = 0;
                }
                for (int i = dataStartIndex; i < lines.Length; i++)
                {
                    var values = ParseCsvLine(lines[i]);
                    if (values.Count > 0)
                    {
                        var rowData = new Dictionary<string, string>();
                        for (int j = 0; j < headers.Count; j++)
                        {
                            string value = j < values.Count ? values[j] : string.Empty;
                            rowData[headers[j]] = value;
                        }
                        result.Add(rowData);
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                throw new Exception($"CSV解析失败: {ex.Message}", ex);
            }
        }
        /// <summary>
        /// 解析CSV行
        /// </summary>
        private List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            if (string.IsNullOrEmpty(line))
                return result;
            bool inQuotes = false;
            var currentField = new StringBuilder();
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        currentField.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(currentField.ToString());
                    currentField.Clear();
                }
                else
                {
                    currentField.Append(c);
                }
            }
            result.Add(currentField.ToString());
            return result;
        }
        /// <summary>
        /// 使用回退机制解码内容
        /// </summary>
        private string DecodeContentWithFallback(byte[] fileContent, EncodingDetectionResult encodingResult)
        {
            try
            {
                if (encodingResult.HasBom)
                {
                    // 跳过BOM字节
                    var bomLength = encodingResult.DetectedEncoding == Encoding.UTF8 ? 3 : 
                                   encodingResult.DetectedEncoding == Encoding.Unicode || encodingResult.DetectedEncoding == Encoding.BigEndianUnicode ? 2 : 4;
                    return encodingResult.DetectedEncoding.GetString(fileContent, bomLength, fileContent.Length - bomLength);
                }
                else
                {
                    return encodingResult.DetectedEncoding.GetString(fileContent);
                }
            }
            catch
            {
                // 如果解码失败，尝试UTF-8
                return Encoding.UTF8.GetString(fileContent);
            }
        }
        /// <summary>
        /// 尝试替代编码
        /// </summary>
        private string TryAlternativeEncodings(byte[] fileContent)
        {
            // 尝试常见的中文编码
            var encodingsToTry = new (string name, Func<Encoding> encodingFactory)[]
            {
                ("GB2312", () => Encoding.GetEncoding("GB2312")),
                ("GBK", () => Encoding.GetEncoding("GBK")),
                ("GB18030", () => Encoding.GetEncoding("GB18030")),
                ("Big5", () => Encoding.GetEncoding("Big5")),
                ("Windows-936", () => Encoding.GetEncoding(936)),
                ("Windows-950", () => Encoding.GetEncoding(950)),
                ("System-Default", () => Encoding.Default)
            };
            string bestResult = null;
            double bestChineseRatio = 0;
            foreach (var (name, encodingFactory) in encodingsToTry)
            {
                try
                {
                    var encoding = encodingFactory();
                    var decoded = encoding.GetString(fileContent);
                    
                    // 检查中文字符比例
                    var chineseRatio = EncodingDetector.CalculateChineseCharacterRatio(decoded);
                    
                    if (chineseRatio > bestChineseRatio)
                    {
                        bestChineseRatio = chineseRatio;
                        bestResult = decoded;
                    }
                }
                catch
                {
                }
            }
            // 如果找到了包含中文字符的结果，使用它
            if (bestChineseRatio > 0 && bestResult != null)
            {
                return bestResult;
            }
            // 否则返回UTF-8解码结果
            return Encoding.UTF8.GetString(fileContent);
        }
        /// <summary>
        /// 解析Excel内容（支持多 Sheet）
        /// </summary>
        private List<Dictionary<string, string>> ParseExcelContent(byte[] fileContent, Type objectType, ExcelImportOptions options)
        {
            try
            {
                using (var stream = new MemoryStream(fileContent))
                {
                    using (var spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
                    {
                        var workbookPart = spreadsheetDocument.WorkbookPart;
                        var sheets = workbookPart.Workbook.Sheets;
                        if (sheets == null || sheets.Count() == 0)
                        {
                            return new List<Dictionary<string, string>>();
                        }
                        // 优先读取第一个 Sheet 作为主表数据
                        var firstSheet = sheets.Elements<Sheet>().FirstOrDefault();
                        if (firstSheet == null)
                        {
                            return new List<Dictionary<string, string>>();
                        }
                        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id);
                        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                        if (sheetData == null)
                        {
                            return new List<Dictionary<string, string>>();
                        }
                        // 获取共享字符串表
                        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
                        return ParseSheetData(sheetData, options, sharedStringTable);
                    }
                }
            }
            catch
            {
                // 回退到 CSV 解析
                try
                {
                    return ParseCsvContent(fileContent, objectType, options);
                }
                catch
                {
                    return new List<Dictionary<string, string>>();
                }
            }
        }
        /// <summary>
        /// 解析 Excel Sheet 数据为字典列表
        /// </summary>
        private List<Dictionary<string, string>> ParseSheetData(SheetData sheetData, ExcelImportOptions options, SharedStringTable sharedStringTable = null)
        {
            var result = new List<Dictionary<string, string>>();
            var rows = sheetData.Elements<Row>().ToList();
            if (rows.Count == 0)
            {
                return result;
            }
            // 直接使用第一行作为表头（不依赖RowIndex）
            var headerRow = rows.FirstOrDefault();
            if (headerRow == null)
            {
                return result;
            }
            // **修复**: 表头也需要按列索引解析，以支持空列
            var headerCells = headerRow.Elements<Cell>().ToList();
            var headerMap = new Dictionary<uint, string>();
            foreach (var cell in headerCells)
            {
                if (cell.CellReference != null)
                {
                    // 解析列索引 (例如 "A1" → 0, "B1" → 1)
                    var columnReference = new string(cell.CellReference.Value.Where(char.IsLetter).ToArray());
                    uint columnIndex = 0;
                    foreach (var c in columnReference)
                    {
                        columnIndex = columnIndex * 26 + (uint)(c - 'A' + 1);
                    }
                    // 转换为从0开始的索引
                    uint zeroBasedIndex = columnIndex - 1;
                    string cellValue = GetCellValue(cell, sharedStringTable);
                    headerMap[zeroBasedIndex] = cellValue;
                }
            }
            // 将映射转换为有序列表
            var maxColumnIndex = headerMap.Keys.Count > 0 ? headerMap.Keys.Max() + 1 : 0;
            var headers = new List<string>();
            for (uint i = 0; i < maxColumnIndex; i++)
            {
                if (headerMap.ContainsKey(i))
                {
                    headers.Add(headerMap[i]);
                }
                else
                {
                    // 空列
                    headers.Add("");
                }
            }
            // 解析数据行（跳过第一行表头）
            var dataRows = rows.Skip(1).ToList();
            foreach (var row in dataRows)
            {
                var rowData = new Dictionary<string, string>();
                var cells = row.Elements<Cell>().ToList();
                // **修复**: 按单元格的实际列索引来获取值,而不是按数组索引
                // 创建列索引到单元格的映射（从0开始）
                var cellMap = new Dictionary<uint, Cell>();
                foreach (var cell in cells)
                {
                    if (cell.CellReference != null)
                    {
                        // 解析列索引 (例如 "A2" → 0, "B2" → 1)
                        var columnReference = new string(cell.CellReference.Value.Where(char.IsLetter).ToArray());
                        uint columnIndex = 0;
                        foreach (var c in columnReference)
                        {
                            columnIndex = columnIndex * 26 + (uint)(c - 'A' + 1);
                        }
                        // 转换为从0开始的索引
                        uint zeroBasedIndex = columnIndex - 1;
                        cellMap[zeroBasedIndex] = cell;
                    }
                }
                for (int i = 0; i < headers.Count; i++)
                {
                    string value = "";
                    // **修复**: 使用从0开始的列索引来获取单元格
                    if (cellMap.ContainsKey((uint)i))
                    {
                        value = GetCellValue(cellMap[(uint)i], sharedStringTable);
                    }
                    if (!string.IsNullOrEmpty(headers[i]))
                    {
                        rowData[headers[i]] = value;
                    }
                }
                // **调试**: 输出每一行的数据
                if (rowData.Count > 0)
                {
                    // 输出所有字段的值(便于诊断)
                    var allFieldsInfo = string.Join(", ", rowData.Keys.Select(k => $"{k}='{rowData[k]}'"));
                    var logMsg = $"[Excel] 解析行数据 ({result.Count + 1}): {allFieldsInfo}";
                    System.Diagnostics.Debug.WriteLine(logMsg);
                    result.Add(rowData);
                }
            }
            var completeMsg = $"[Excel] 解析完成，返回 {result.Count} 行数据";
            return result;
        }
        /// <summary>
        /// 获取单元格的值
        /// </summary>
        private string GetCellValue(Cell cell, SharedStringTable sharedStringTable = null)
        {
            if (cell == null || cell.CellValue == null)
            {
                return "";
            }
            string value = cell.CellValue.InnerText;
            // 如果是共享字符串表，需要从共享字符串表中获取实际值
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (sharedStringTable != null)
                {
                    try
                    {
                        int index = int.Parse(value);
                        var sharedStringItems = sharedStringTable.Elements<SharedStringItem>().ToList();
                        if (index >= 0 && index < sharedStringItems.Count)
                        {
                            value = sharedStringItems[index].InnerText;
                        }
                        else
                        {
                            value = "";
                        }
                    }
                    catch
                    {
                        value = "";
                    }
                }
                else
                {
                }
            }
            return value;
        }
        /// <summary>
        /// 解析所有 Sheet 数据（用于主子表导入）
        /// </summary>
        private List<MultiSheetData> ParseAllSheets(byte[] fileContent, ExcelImportOptions options)
        {
            var result = new List<MultiSheetData>();
            try
            {
                using (var stream = new MemoryStream(fileContent))
                {
                    using (var spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
                    {
                        var workbookPart = spreadsheetDocument.WorkbookPart;
                        var sheets = workbookPart.Workbook.Sheets;
                        if (sheets == null || sheets.Count() == 0)
                        {
                            return result;
                        }
                        // 获取共享字符串表（所有Sheet共享同一个）
                        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
                        foreach (var sheet in sheets.Elements<Sheet>())
                        {
                            try
                            {
                                var worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
                                if (worksheetPart == null)
                                {
                                    continue;
                                }
                                var worksheet = worksheetPart.Worksheet;
                                if (worksheet == null)
                                {
                                    continue;
                                }
                                var sheetData = worksheet.GetFirstChild<SheetData>();
                                if (sheetData != null)
                                {
                                    var dataRows = ParseSheetData(sheetData, options, sharedStringTable);
                                    result.Add(new MultiSheetData
                                    {
                                        SheetName = sheet.Name.Value,
                                        DataRows = dataRows
                                    });
                                }
                                else
                                {
                                }
                            }
                            catch
                            {
                            }
                        }
                    }
                }
            }
            catch
            {
            }
            return result;
        }
        /// <summary>
        /// 导出多个数据集到单个 Excel 文件的多个 Sheet (仅支持 XLSX 格式)
        /// </summary>
        public async Task ExportMultipleSheetsAsync<T>(
            string filePath,
            Dictionary<string, List<T>> sheetsData,
            string[] headers,
            Func<T, object[]> dataExtractor,
            string title = null)
        {
            await Task.Run(() =>
            {
                using (var spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = spreadsheetDocument.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                    uint sheetId = 1;
                    foreach (var sheetData in sheetsData)
                    {
                        var sheetName = sheetData.Key;
                        var data = sheetData.Value;
                        // 创建工作表
                        var worksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet(new SheetData());
                        var sheet = new Sheet()
                        {
                            Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                            SheetId = sheetId++,
                            Name = sheetName
                        };
                        sheets.Append(sheet);
                        var sheetDataElement = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                        // 添加表头
                        var headerRow = new Row();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            var cell = CreateTextCell((uint)(i + 1), 1, headers[i]);
                            headerRow.AppendChild(cell);
                        }
                        sheetDataElement.AppendChild(headerRow);
                        // 添加数据行
                        uint rowIndex = 2;
                        foreach (var item in data)
                        {
                            var dataRow = new Row() { RowIndex = rowIndex++ };
                            var cellValues = dataExtractor(item);
                            for (int i = 0; i < cellValues.Length && i < headers.Length; i++)
                            {
                                var cellValue = cellValues[i]?.ToString() ?? "";
                                var cell = CreateTextCell((uint)(i + 1), rowIndex - 1, cellValue);
                                dataRow.AppendChild(cell);
                            }
                            sheetDataElement.AppendChild(dataRow);
                        }
                    }
                    spreadsheetDocument.WorkbookPart.Workbook.Save();
                }
            });
        }
        /// <summary>
        /// 获取 Excel 文件中所有 Sheet 的数据
        /// </summary>
        public async Task<Dictionary<string, List<Dictionary<string, object>>>> GetExcelSheetsAsync(string filePath)
        {
            return await Task.Run(() =>
            {
                var result = new Dictionary<string, List<Dictionary<string, object>>>();
                try
                {
                    using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
                        {
                            var workbookPart = spreadsheetDocument.WorkbookPart;
                            var sheets = workbookPart.Workbook.Sheets;
                            if (sheets == null || sheets.Count() == 0)
                            {
                                return result;
                            }
                            // 获取共享字符串表
                            var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
                            foreach (var sheet in sheets.Elements<Sheet>())
                            {
                                try
                                {
                                    var worksheetPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
                                    if (worksheetPart == null) continue;
                                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                                    if (sheetData == null) continue;
                                    var dataRows = ParseSheetDataToObject(sheetData, sharedStringTable);
                                    result[sheet.Name.Value] = dataRows;
                                }
                                catch
                                {
                                }
                            }
                        }
                    }
                }
                catch
                {
                    throw;
                }
                return result;
            });
        }
        /// <summary>
        /// 解析 Excel Sheet 数据为对象字典列表
        /// </summary>
        private List<Dictionary<string, object>> ParseSheetDataToObject(SheetData sheetData, SharedStringTable sharedStringTable = null)
        {
            var result = new List<Dictionary<string, object>>();
            var rows = sheetData.Elements<Row>().ToList();
            if (rows.Count == 0)
            {
                return result;
            }
            // 使用第一行作为表头
            var headerRow = rows.FirstOrDefault();
            if (headerRow == null)
            {
                return result;
            }
            var headers = new List<string>();
            foreach (var cell in headerRow.Elements<Cell>())
            {
                string cellValue = GetCellValue(cell, sharedStringTable);
                headers.Add(cellValue);
            }
            // 解析数据行
            var dataRows = rows.Skip(1).ToList();
            foreach (var row in dataRows)
            {
                var rowData = new Dictionary<string, object>();
                var cells = row.Elements<Cell>().ToList();
                for (int i = 0; i < headers.Count; i++)
                {
                    object value = "";
                    if (i < cells.Count)
                    {
                        value = GetCellValue(cells[i], sharedStringTable);
                    }
                    if (!string.IsNullOrEmpty(headers[i]))
                    {
                        rowData[headers[i]] = value;
                    }
                }
                if (rowData.Count > 0)
                {
                    result.Add(rowData);
                }
            }
            return result;
        }
        #endregion
        #region 内部类
        /// <summary>
        /// 多Sheet数据容器
        /// </summary>
        private class MultiSheetData
        {
            public string SheetName { get; set; }
            public List<Dictionary<string, string>> DataRows { get; set; } = new List<Dictionary<string, string>>();
        }
        /// <summary>
        /// 文件解析结果（支持多 Sheet）
        /// </summary>
        private class FileParseResult
        {
            public bool IsSuccess { get; set; }
            public string ErrorMessage { get; set; }
            public List<Dictionary<string, string>> DataRows { get; set; } = new List<Dictionary<string, string>>();
            // 新增多 Sheet 支持
            public List<MultiSheetData> Sheets { get; set; } = new List<MultiSheetData>();
            public bool HasMultipleSheets => Sheets.Count > 1;
        }
        /// <summary>
        /// 导入行结果
        /// </summary>
        private class ImportRowResult
        {
            public bool IsSuccess { get; set; }
            public List<ExcelImportError> Errors { get; set; } = new List<ExcelImportError>();
            public List<ExcelImportWarning> Warnings { get; set; } = new List<ExcelImportWarning>();
            public object CreatedObject { get; set; } // 新增：保存创建的对象
        }
        /// <summary>
        /// 值转换结果
        /// </summary>
        private class ValueConvertResult
        {
            public bool IsSuccess { get; set; }
            public object Value { get; set; }
            public string ErrorMessage { get; set; }
            public bool HasWarning { get; set; }
            public string WarningMessage { get; set; }
        }
        /// <summary>
        /// 解析枚举映射格式
        /// </summary>
        /// <param name="format">格式字符串，如 "Male=男;Female=女"</param>
        /// <returns>映射字典</returns>
        private Dictionary<string, string> ParseEnumMappings(string format)
        {
            var mappings = new Dictionary<string, string>();
            try
            {
                var pairs = format.Split(';');
                foreach (var pair in pairs)
                {
                    var parts = pair.Split('=');
                    if (parts.Length == 2)
                    {
                        mappings[parts[0].Trim()] = parts[1].Trim();
                    }
                }
            }
            catch
            {
                // 忽略解析错误
            }
            return mappings;
        }
        #endregion
    }
}