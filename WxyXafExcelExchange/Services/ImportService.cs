using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Wxy.Xaf.ExcelExchange;

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// Excel导入服务实现
    /// </summary>
    public class ImportService : IImportService
    {
        /// <summary>
        /// 从CSV数据导入对象列表
        /// </summary>
        /// <typeparam name="T">目标对象类型</typeparam>
        /// <param name="csvData">CSV数据</param>
        /// <param name="options">导入选项</param>
        /// <returns>导入结果</returns>
        public async Task<ImportResult<T>> ImportFromCsvAsync<T>(string csvData, ImportOptions options = null) where T : class, new()
        {
            var result = await ImportFromCsvAsync(typeof(T), csvData, options);
            
            return new ImportResult<T>
            {
                IsSuccess = result.IsSuccess,
                TotalRecords = result.TotalRecords,
                SuccessfulRecords = result.SuccessfulRecords,
                FailedRecords = result.FailedRecords,
                Errors = result.Errors,
                ProcessingTimeMs = result.ProcessingTimeMs,
                SuccessfulObjects = result.SuccessfulObjects.Cast<T>().ToList()
            };
        }

        /// <summary>
        /// 从CSV数据导入对象列表（非泛型版本）
        /// </summary>
        /// <param name="objectType">目标对象类型</param>
        /// <param name="csvData">CSV数据</param>
        /// <param name="options">导入选项</param>
        /// <returns>导入结果</returns>
        public async Task<ImportResult> ImportFromCsvAsync(Type objectType, string csvData, ImportOptions options = null)
        {
            var startTime = DateTime.UtcNow;
            options = options ?? new ImportOptions();
            
            var result = new ImportResult();
            
            try
            {
                // 解析CSV数据
                var csvRows = ParseCsvData(csvData, options);
                if (csvRows.Count == 0)
                {
                    result.AddError(new ImportError
                    {
                        RowIndex = 0,
                        Message = "CSV数据为空或格式无效",
                        ErrorType = ImportErrorType.ParseError
                    });
                    return result;
                }

                // 获取可导入的属性
                var importableProperties = AttributeHelper.GetImportableProperties(objectType);
                var columnToPropertyMap = AttributeHelper.CreateColumnToPropertyMap(importableProperties);

                // 获取表头
                var headers = csvRows[0];
                var dataRows = csvRows.Skip(options.HasHeaderRow ? 1 : 0).ToList();
                
                result.TotalRecords = dataRows.Count;

                // 验证表头
                var validationResult = ValidateHeaders(objectType, headers, importableProperties);
                if (!validationResult.IsValid)
                {
                    foreach (var error in validationResult.Errors)
                    {
                        result.AddError(new ImportError
                        {
                            RowIndex = 0,
                            Message = error,
                            ErrorType = ImportErrorType.ValidationError
                        });
                    }
                }

                // 处理数据行
                var rowIndex = options.HasHeaderRow ? 2 : 1; // 从第2行开始（如果有表头）
                var processedCount = 0;

                foreach (var dataRow in dataRows)
                {
                    try
                    {
                        // 报告进度
                        if (options.EnableProgressReporting && options.ProgressCallback != null)
                        {
                            options.ProgressCallback(new ImportProgress
                            {
                                CurrentRow = processedCount + 1,
                                TotalRows = dataRows.Count,
                                Stage = "处理数据行",
                                SuccessfulRecords = result.SuccessfulRecords,
                                FailedRecords = result.FailedRecords
                            });
                        }

                        var obj = await ProcessDataRowAsync(objectType, headers, dataRow, columnToPropertyMap, rowIndex, result, options);
                        if (obj != null)
                        {
                            result.SuccessfulObjects.Add(obj);
                            result.SuccessfulRecords++;
                        }
                        else
                        {
                            result.FailedRecords++;
                        }
                    }
                    catch (Exception ex)
                    {
                        result.AddError(new ImportError
                        {
                            RowIndex = rowIndex,
                            Message = $"处理行时发生错误: {ex.Message}",
                            ErrorType = ImportErrorType.SystemError,
                            Exception = ex
                        });
                        result.FailedRecords++;
                    }

                    rowIndex++;
                    processedCount++;
                }

                result.IsSuccess = result.FailedRecords == 0 || (options.SkipValidationErrors && result.SuccessfulRecords > 0);
            }
            catch (Exception ex)
            {
                result.AddError(new ImportError
                {
                    RowIndex = 0,
                    Message = $"导入过程中发生系统错误: {ex.Message}",
                    ErrorType = ImportErrorType.SystemError,
                    Exception = ex
                });
                result.IsSuccess = false;
            }

            result.ProcessingTimeMs = (long)(DateTime.UtcNow - startTime).TotalMilliseconds;
            return result;
        }

        /// <summary>
        /// 验证CSV数据格式
        /// </summary>
        /// <param name="objectType">目标对象类型</param>
        /// <param name="csvData">CSV数据</param>
        /// <returns>验证结果</returns>
        public HeaderValidationResult ValidateCsvFormat(Type objectType, string csvData)
        {
            var result = new HeaderValidationResult { IsValid = true };

            try
            {
                // 解析CSV数据
                var csvRows = ParseCsvData(csvData, new ImportOptions());
                if (csvRows.Count == 0)
                {
                    result.IsValid = false;
                    result.Errors.Add("CSV数据为空");
                    return result;
                }

                // 获取表头和数据行数
                var headers = csvRows[0];
                result.DetectedColumns = headers.ToList();
                result.DataRowCount = csvRows.Count - 1;

                // 获取可导入的属性
                var importableProperties = AttributeHelper.GetImportableProperties(objectType);
                var requiredColumns = importableProperties
                    .Where(p => AttributeHelper.IsRequired(p))
                    .Select(p => AttributeHelper.GetEffectiveColumnName(p))
                    .ToList();

                // 检查必需列
                foreach (var requiredColumn in requiredColumns)
                {
                    if (!headers.Contains(requiredColumn, StringComparer.OrdinalIgnoreCase))
                    {
                        result.MissingRequiredColumns.Add(requiredColumn);
                        result.IsValid = false;
                    }
                }

                // 检查未识别的列
                var recognizedColumns = importableProperties
                    .Select(p => AttributeHelper.GetEffectiveColumnName(p))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                foreach (var header in headers)
                {
                    if (!recognizedColumns.Contains(header))
                    {
                        result.UnrecognizedColumns.Add(header);
                        result.Warnings.Add($"未识别的列: {header}");
                    }
                }

                // 添加错误消息
                if (result.MissingRequiredColumns.Any())
                {
                    result.Errors.Add($"缺少必需的列: {string.Join(", ", result.MissingRequiredColumns)}");
                }

                if (result.DataRowCount == 0)
                {
                    result.Warnings.Add("没有数据行");
                }
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.Errors.Add($"CSV格式验证失败: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 获取对象类型的导入模板（CSV表头）
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <returns>CSV表头字符串</returns>
        public string GetImportTemplate(Type objectType)
        {
            var importableProperties = AttributeHelper.GetImportableProperties(objectType);
            var columnNames = importableProperties
                .Select(p => AttributeHelper.GetEffectiveColumnName(p))
                .ToList();

            return string.Join(",", columnNames.Select(name => $"\"{name}\""));
        }

        /// <summary>
        /// 获取对象类型的可导入字段信息
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <returns>字段信息列表</returns>
        public List<ImportFieldInfo> GetImportableFields(Type objectType)
        {
            var importableProperties = AttributeHelper.GetImportableProperties(objectType);
            
            return importableProperties.Select(p => new ImportFieldInfo
            {
                PropertyName = p.Name,
                ColumnName = AttributeHelper.GetEffectiveColumnName(p),
                PropertyType = p.PropertyType,
                IsRequired = AttributeHelper.IsRequired(p),
                DefaultValue = AttributeHelper.GetDefaultValue(p),
                Format = AttributeHelper.GetCustomFormat(p),
                LookupKey = AttributeHelper.GetLookupKey(p),
                Description = GetPropertyDescription(p)
            }).ToList();
        }

        #region 私有方法

        /// <summary>
        /// 解析CSV数据
        /// </summary>
        /// <param name="csvData">CSV数据</param>
        /// <param name="options">导入选项</param>
        /// <returns>CSV行列表</returns>
        private List<List<string>> ParseCsvData(string csvData, ImportOptions options)
        {
            var rows = new List<List<string>>();
            
            if (string.IsNullOrWhiteSpace(csvData))
                return rows;

            using (var reader = new StringReader(csvData))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    var columns = ParseCsvLine(line, options.CsvSeparator);
                    rows.Add(columns);
                }
            }

            return rows;
        }

        /// <summary>
        /// 解析CSV行
        /// </summary>
        /// <param name="line">CSV行</param>
        /// <param name="separator">分隔符</param>
        /// <returns>列值列表</returns>
        private List<string> ParseCsvLine(string line, char separator)
        {
            var columns = new List<string>();
            var currentColumn = new StringBuilder();
            var inQuotes = false;
            var i = 0;

            while (i < line.Length)
            {
                var c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        // 转义的引号
                        currentColumn.Append('"');
                        i += 2;
                    }
                    else
                    {
                        // 切换引号状态
                        inQuotes = !inQuotes;
                        i++;
                    }
                }
                else if (c == separator && !inQuotes)
                {
                    // 列分隔符
                    columns.Add(currentColumn.ToString().Trim());
                    currentColumn.Clear();
                    i++;
                }
                else
                {
                    currentColumn.Append(c);
                    i++;
                }
            }

            // 添加最后一列
            columns.Add(currentColumn.ToString().Trim());

            return columns;
        }

        /// <summary>
        /// 验证表头
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <param name="headers">表头列表</param>
        /// <param name="importableProperties">可导入属性列表</param>
        /// <returns>验证结果</returns>
        private HeaderValidationResult ValidateHeaders(Type objectType, List<string> headers, List<PropertyInfo> importableProperties)
        {
            var result = new HeaderValidationResult { IsValid = true };

            // 检查必需字段
            var requiredColumns = importableProperties
                .Where(p => AttributeHelper.IsRequired(p))
                .Select(p => AttributeHelper.GetEffectiveColumnName(p))
                .ToList();

            foreach (var requiredColumn in requiredColumns)
            {
                if (!headers.Contains(requiredColumn, StringComparer.OrdinalIgnoreCase))
                {
                    result.IsValid = false;
                    result.Errors.Add($"缺少必需的列: {requiredColumn}");
                }
            }

            return result;
        }

        /// <summary>
        /// 处理数据行 - 支持导入模式
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <param name="headers">表头列表</param>
        /// <param name="dataRow">数据行</param>
        /// <param name="columnToPropertyMap">列到属性的映射</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="result">导入结果</param>
        /// <param name="options">导入选项</param>
        /// <returns>创建的对象</returns>
        private async Task<object> ProcessDataRowAsync(Type objectType, List<string> headers, List<string> dataRow, 
            Dictionary<string, PropertyInfo> columnToPropertyMap, int rowIndex, ImportResult result, ImportOptions options)
        {
            try
            {
                // 根据导入模式处理记录
                switch (options.ImportMode)
                {
                    case ImportMode.CreateOnly:
                        return await ProcessCreateOnlyAsync(objectType, headers, dataRow, columnToPropertyMap, rowIndex, result, options);
                    
                    case ImportMode.UpdateOnly:
                        return await ProcessUpdateOnlyAsync(objectType, headers, dataRow, columnToPropertyMap, rowIndex, result, options);
                    
                    case ImportMode.CreateOrUpdate:
                        return await ProcessCreateOrUpdateAsync(objectType, headers, dataRow, columnToPropertyMap, rowIndex, result, options);
                    
                    case ImportMode.ReplaceAll:
                        // ReplaceAll模式在导入开始时处理删除，这里只创建新记录
                        return await ProcessCreateOnlyAsync(objectType, headers, dataRow, columnToPropertyMap, rowIndex, result, options);
                    
                    default:
                        return await ProcessCreateOrUpdateAsync(objectType, headers, dataRow, columnToPropertyMap, rowIndex, result, options);
                }
            }
            catch (Exception ex)
            {
                result.AddError(rowIndex, "", $"处理数据行时发生错误: {ex.Message}", ImportErrorType.SystemError);
                return null;
            }
        }

        /// <summary>
        /// 处理仅创建模式
        /// </summary>
        private async Task<object> ProcessCreateOnlyAsync(Type objectType, List<string> headers, List<string> dataRow, 
            Dictionary<string, PropertyInfo> columnToPropertyMap, int rowIndex, ImportResult result, ImportOptions options)
        {
            // 检查是否存在重复记录
            if (await IsDuplicateRecordAsync(objectType, headers, dataRow, columnToPropertyMap, options))
            {
                result.AddError(rowIndex, "", "记录已存在，跳过创建", ImportErrorType.ValidationError);
                return null;
            }

            return await CreateNewObjectAsync(objectType, headers, dataRow, columnToPropertyMap, rowIndex, result, options);
        }

        /// <summary>
        /// 处理仅更新模式
        /// </summary>
        private async Task<object> ProcessUpdateOnlyAsync(Type objectType, List<string> headers, List<string> dataRow, 
            Dictionary<string, PropertyInfo> columnToPropertyMap, int rowIndex, ImportResult result, ImportOptions options)
        {
            // 查找现有记录
            var existingObject = await FindExistingRecordAsync(objectType, headers, dataRow, columnToPropertyMap, options);
            if (existingObject == null)
            {
                result.AddError(rowIndex, "", "未找到要更新的记录，跳过", ImportErrorType.ValidationError);
                return null;
            }

            return await UpdateExistingObjectAsync(existingObject, headers, dataRow, columnToPropertyMap, rowIndex, result, options);
        }

        /// <summary>
        /// 处理创建或更新模式
        /// </summary>
        private async Task<object> ProcessCreateOrUpdateAsync(Type objectType, List<string> headers, List<string> dataRow, 
            Dictionary<string, PropertyInfo> columnToPropertyMap, int rowIndex, ImportResult result, ImportOptions options)
        {
            // 查找现有记录
            var existingObject = await FindExistingRecordAsync(objectType, headers, dataRow, columnToPropertyMap, options);
            
            if (existingObject != null)
            {
                // 更新现有记录
                return await UpdateExistingObjectAsync(existingObject, headers, dataRow, columnToPropertyMap, rowIndex, result, options);
            }
            else
            {
                // 创建新记录
                return await CreateNewObjectAsync(objectType, headers, dataRow, columnToPropertyMap, rowIndex, result, options);
            }
        }

        /// <summary>
        /// 创建新对象
        /// </summary>
        private async Task<object> CreateNewObjectAsync(Type objectType, List<string> headers, List<string> dataRow, 
            Dictionary<string, PropertyInfo> columnToPropertyMap, int rowIndex, ImportResult result, ImportOptions options)
        {
            var obj = Activator.CreateInstance(objectType);
            var hasErrors = false;

            for (int i = 0; i < Math.Min(headers.Count, dataRow.Count); i++)
            {
                var columnName = headers[i];
                var cellValue = dataRow[i];

                if (!columnToPropertyMap.TryGetValue(columnName, out var property))
                    continue; // 跳过未识别的列

                try
                {
                    // 获取最终值（考虑默认值）
                    var finalValue = AttributeHelper.GetFinalValue(property, cellValue);

                    // 验证属性值
                    if (!AttributeHelper.ValidatePropertyValue(property, finalValue, out var errorMessage))
                    {
                        result.AddError(rowIndex, columnName, errorMessage, ImportErrorType.ValidationError);
                        hasErrors = true;
                        continue;
                    }

                    // 设置属性值
                    await SetPropertyValueAsync(obj, property, finalValue, rowIndex, columnName, result, options);
                }
                catch (Exception ex)
                {
                    result.AddError(rowIndex, columnName, $"设置属性值时发生错误: {ex.Message}", ImportErrorType.SystemError);
                    hasErrors = true;
                }
            }

            // 如果有错误且不跳过验证错误，返回null
            if (hasErrors && !options.SkipValidationErrors)
                return null;

            return obj;
        }

        /// <summary>
        /// 更新现有对象
        /// </summary>
        private async Task<object> UpdateExistingObjectAsync(object existingObject, List<string> headers, List<string> dataRow, 
            Dictionary<string, PropertyInfo> columnToPropertyMap, int rowIndex, ImportResult result, ImportOptions options)
        {
            var hasErrors = false;

            for (int i = 0; i < Math.Min(headers.Count, dataRow.Count); i++)
            {
                var columnName = headers[i];
                var cellValue = dataRow[i];

                if (!columnToPropertyMap.TryGetValue(columnName, out var property))
                    continue; // 跳过未识别的列

                // 跳过重复检测键属性（不更新用于查找的键）
                if (options.DuplicateDetectionKeys.Contains(property.Name))
                    continue;

                try
                {
                    // 获取最终值（考虑默认值）
                    var finalValue = AttributeHelper.GetFinalValue(property, cellValue);

                    // 验证属性值
                    if (!AttributeHelper.ValidatePropertyValue(property, finalValue, out var errorMessage))
                    {
                        result.AddError(rowIndex, columnName, errorMessage, ImportErrorType.ValidationError);
                        hasErrors = true;
                        continue;
                    }

                    // 设置属性值
                    await SetPropertyValueAsync(existingObject, property, finalValue, rowIndex, columnName, result, options);
                }
                catch (Exception ex)
                {
                    result.AddError(rowIndex, columnName, $"更新属性值时发生错误: {ex.Message}", ImportErrorType.SystemError);
                    hasErrors = true;
                }
            }

            // 如果有错误且不跳过验证错误，返回null
            if (hasErrors && !options.SkipValidationErrors)
                return null;

            return existingObject;
        }

        /// <summary>
        /// 检查是否为重复记录
        /// </summary>
        private async Task<bool> IsDuplicateRecordAsync(Type objectType, List<string> headers, List<string> dataRow, 
            Dictionary<string, PropertyInfo> columnToPropertyMap, ImportOptions options)
        {
            var existingRecord = await FindExistingRecordAsync(objectType, headers, dataRow, columnToPropertyMap, options);
            return existingRecord != null;
        }

        /// <summary>
        /// 查找现有记录
        /// </summary>
        private async Task<object> FindExistingRecordAsync(Type objectType, List<string> headers, List<string> dataRow, 
            Dictionary<string, PropertyInfo> columnToPropertyMap, ImportOptions options)
        {
            // 如果没有指定重复检测键，使用所有属性进行匹配（简单实现）
            if (options.DuplicateDetectionKeys == null || options.DuplicateDetectionKeys.Count == 0)
            {
                // 简单实现：总是返回null，表示没有找到重复记录
                // 在实际应用中，这里应该连接到数据访问层进行查询
                return null;
            }

            // 构建查询条件
            var queryConditions = new Dictionary<string, object>();
            
            for (int i = 0; i < Math.Min(headers.Count, dataRow.Count); i++)
            {
                var columnName = headers[i];
                var cellValue = dataRow[i];

                if (!columnToPropertyMap.TryGetValue(columnName, out var property))
                    continue;

                if (options.DuplicateDetectionKeys.Contains(property.Name))
                {
                    try
                    {
                        var finalValue = AttributeHelper.GetFinalValue(property, cellValue);
                        var convertedValue = ConvertValue(finalValue, property.PropertyType, AttributeHelper.GetCustomFormat(property));
                        queryConditions[property.Name] = convertedValue;
                    }
                    catch
                    {
                        // 如果转换失败，跳过这个条件
                        continue;
                    }
                }
            }

            // 这里应该实现实际的数据库查询逻辑
            // 由于ImportService是独立的，没有直接的数据访问能力，
            // 在实际使用中需要注入数据访问服务或ObjectSpace
            
            // 占位符实现：总是返回null
            await Task.CompletedTask; // 保持异步签名
            return null;
        }

        /// <summary>
        /// 设置属性值
        /// </summary>
        /// <param name="obj">目标对象</param>
        /// <param name="property">属性信息</param>
        /// <param name="value">属性值</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnName">列名</param>
        /// <param name="result">导入结果</param>
        /// <param name="options">导入选项</param>
        private async Task SetPropertyValueAsync(object obj, PropertyInfo property, string value, 
            int rowIndex, string columnName, ImportResult result, ImportOptions options)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                // 处理空值
                if (property.PropertyType.IsValueType && Nullable.GetUnderlyingType(property.PropertyType) == null)
                {
                    // 非可空值类型，使用默认值
                    var defaultValue = Activator.CreateInstance(property.PropertyType);
                    property.SetValue(obj, defaultValue);
                }
                else
                {
                    // 可空类型或引用类型，设置为null
                    property.SetValue(obj, null);
                }
                return;
            }

            try
            {
                // 检查是否有关联对象查找
                if (AttributeHelper.HasLookupKey(property))
                {
                    await SetLookupPropertyValueAsync(obj, property, value, rowIndex, columnName, result, options);
                    return;
                }

                // 获取自定义格式
                var customFormat = AttributeHelper.GetCustomFormat(property);
                
                // 类型转换
                var convertedValue = ConvertValue(value, property.PropertyType, customFormat);
                property.SetValue(obj, convertedValue);
            }
            catch (Exception ex)
            {
                result.AddError(rowIndex, columnName, $"类型转换失败: {ex.Message}", ImportErrorType.ConversionError);
                throw;
            }
        }

        /// <summary>
        /// 设置关联对象属性值
        /// </summary>
        /// <param name="obj">目标对象</param>
        /// <param name="property">属性信息</param>
        /// <param name="lookupValue">查找值</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnName">列名</param>
        /// <param name="result">导入结果</param>
        /// <param name="options">导入选项</param>
        private async Task SetLookupPropertyValueAsync(object obj, PropertyInfo property, string lookupValue,
            int rowIndex, string columnName, ImportResult result, ImportOptions options)
        {
            var lookupKey = AttributeHelper.GetLookupKey(property);
            
            // 这里应该实现实际的对象查找逻辑
            // 由于没有具体的数据访问层，这里只是一个占位符实现
            // 在实际使用中，需要注入数据访问服务来执行查找
            
            // 占位符实现：创建一个简单的对象
            if (options.CreateMissingReferences)
            {
                try
                {
                    var referencedObject = Activator.CreateInstance(property.PropertyType);
                    
                    // 尝试设置查找键属性
                    var keyProperty = property.PropertyType.GetProperty(lookupKey);
                    if (keyProperty != null && keyProperty.CanWrite)
                    {
                        var convertedValue = ConvertValue(lookupValue, keyProperty.PropertyType, null);
                        keyProperty.SetValue(referencedObject, convertedValue);
                    }
                    
                    property.SetValue(obj, referencedObject);
                }
                catch (Exception ex)
                {
                    result.AddError(rowIndex, columnName, $"创建关联对象失败: {ex.Message}", ImportErrorType.LookupError);
                    throw;
                }
            }
            else
            {
                result.AddError(rowIndex, columnName, $"找不到关联对象，查找键: {lookupKey}={lookupValue}", ImportErrorType.LookupError);
                throw new InvalidOperationException($"找不到关联对象: {lookupKey}={lookupValue}");
            }
        }

        /// <summary>
        /// 转换值类型 - 改进空值处理逻辑
        /// </summary>
        /// <param name="value">原始值</param>
        /// <param name="targetType">目标类型</param>
        /// <param name="customFormat">自定义格式</param>
        /// <returns>转换后的值</returns>
        private object ConvertValue(string value, Type targetType, string customFormat)
        {
            // 保存原始目标类型以便后续检查
            var originalTargetType = targetType;
            
            // 统一的空值处理逻辑
            if (string.IsNullOrWhiteSpace(value))
            {
                return GetDefaultValueForType(originalTargetType);
            }

            // 处理可空类型
            var underlyingType = Nullable.GetUnderlyingType(targetType);
            if (underlyingType != null)
                targetType = underlyingType;

            // 使用自定义格式解析
            if (!string.IsNullOrEmpty(customFormat))
            {
                var formatResult = FormatHelper.ParseWithCustomFormat(value, targetType, customFormat);
                if (formatResult.Success)
                    return formatResult.Value;
            }

            // 标准类型转换
            if (targetType == typeof(string))
                return value;

            if (targetType == typeof(int))
                return int.Parse(FormatHelper.CleanNumericValue(value), CultureInfo.InvariantCulture);

            if (targetType == typeof(long))
                return long.Parse(FormatHelper.CleanNumericValue(value), CultureInfo.InvariantCulture);

            if (targetType == typeof(decimal))
                return decimal.Parse(FormatHelper.CleanNumericValue(value), CultureInfo.InvariantCulture);

            if (targetType == typeof(double))
                return double.Parse(FormatHelper.CleanNumericValue(value), CultureInfo.InvariantCulture);

            if (targetType == typeof(float))
                return float.Parse(FormatHelper.CleanNumericValue(value), CultureInfo.InvariantCulture);

            if (targetType == typeof(bool))
            {
                var lowerValue = value.ToLower();
                return lowerValue == "true" || lowerValue == "1" || lowerValue == "是" || lowerValue == "yes";
            }

            if (targetType == typeof(DateTime))
            {
                // 尝试解析DateTime
                if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateResult))
                    return dateResult;
                
                // 解析失败，抛出异常
                throw new FormatException($"无法将 '{value}' 转换为DateTime类型");
            }

            if (targetType.IsEnum)
            {
                // 对值进行 trim，去除前后空格
                var trimmedValue = value.Trim();

                // 如果有自定义格式且包含枚举映射（如 "Male=男;Female=女"）
                if (!string.IsNullOrEmpty(customFormat) && customFormat.Contains("="))
                {
                    var mappings = ParseEnumMappings(customFormat);
                    // 创建反向映射：显示值 -> 枚举名
                    var reverseMappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var mapping in mappings)
                    {
                        reverseMappings[mapping.Value] = mapping.Key;
                    }

                    // 尝试通过显示值查找枚举名
                    if (reverseMappings.TryGetValue(trimmedValue, out string enumName))
                    {
                        if (Enum.IsDefined(targetType, enumName))
                        {
                            return Enum.Parse(targetType, enumName, true);
                        }
                    }
                }

                // 直接解析枚举
                return Enum.Parse(targetType, trimmedValue, true);
            }

            // 尝试使用Convert.ChangeType
            return Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// 获取类型的默认值 - 统一的空值处理逻辑
        /// </summary>
        /// <param name="type">类型</param>
        /// <returns>默认值</returns>
        private object GetDefaultValueForType(Type type)
        {
            // 对于可空类型，返回null
            if (Nullable.GetUnderlyingType(type) != null)
                return null;
            
            // 对于字符串类型，返回空字符串而不是null（业务逻辑考虑）
            if (type == typeof(string))
                return string.Empty;
            
            // 对于值类型，返回默认值
            if (type.IsValueType)
                return Activator.CreateInstance(type);
            
            // 对于引用类型，返回null
            return null;
        }

        /// <summary>
        /// 获取属性描述
        /// </summary>
        /// <param name="property">属性信息</param>
        /// <returns>属性描述</returns>
        private string GetPropertyDescription(PropertyInfo property)
        {
            var excelField = AttributeHelper.GetExcelFieldAttribute(property);
            if (excelField != null)
            {
                var parts = new List<string>();
                
                if (excelField.Required)
                    parts.Add("必填");
                
                if (excelField.HasDefaultValue())
                    parts.Add($"默认值: {excelField.DefaultValue}");
                
                if (excelField.HasCustomFormat())
                    parts.Add($"格式: {excelField.Format}");
                
                if (excelField.HasKeyLookup())
                    parts.Add($"关联查找: {excelField.Key}");

                if (parts.Any())
                    return string.Join(", ", parts);
            }

            return $"{property.PropertyType.Name}类型";
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

    /// <summary>
    /// 表头验证结果
    /// </summary>
    public class HeaderValidationResult
    {
        /// <summary>
        /// 是否验证通过
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// 错误消息列表
        /// </summary>
        public List<string> Errors { get; set; } = new List<string>();

        /// <summary>
        /// 检测到的列名
        /// </summary>
        public List<string> DetectedColumns { get; set; } = new List<string>();

        /// <summary>
        /// 数据行数
        /// </summary>
        public int DataRowCount { get; set; }

        /// <summary>
        /// 缺少的必需列
        /// </summary>
        public List<string> MissingRequiredColumns { get; set; } = new List<string>();

        /// <summary>
        /// 无法识别的列
        /// </summary>
        public List<string> UnrecognizedColumns { get; set; } = new List<string>();

        /// <summary>
        /// 警告消息列表
        /// </summary>
        public List<string> Warnings { get; set; } = new List<string>();
    }
}