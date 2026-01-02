using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Wxy.Xaf.ExcelExchange.Configuration;
using Wxy.Xaf.ExcelExchange.Services;

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// 数据验证服务，负责根据字段配置验证数据
    /// </summary>
    public class DataValidator
    {
        private readonly ConcurrentDictionary<string, Regex> _regexCache = new ConcurrentDictionary<string, Regex>();

        /// <summary>
        /// 验证字段值
        /// </summary>
        /// <param name="value">字段值</param>
        /// <param name="fieldConfig">字段配置</param>
        /// <param name="rowNumber">行号（用于错误报告）</param>
        /// <returns>验证结果</returns>
        public ValidationResult ValidateField(string value, ExcelFieldConfiguration fieldConfig, int rowNumber)
        {
            var result = new ValidationResult
            {
                IsValid = true,
                FieldName = fieldConfig.PropertyInfo.Name,
                RowNumber = rowNumber,
                OriginalValue = value
            };

            try
            {
                // 必填验证
                if (fieldConfig.FieldAttribute.IsRequired && string.IsNullOrWhiteSpace(value))
                {
                    result.IsValid = false;
                    result.ErrorMessage = $"必填字段不能为空";
                    result.ErrorType = ExcelImportErrorType.RequiredFieldEmpty;
                    return result;
                }

                // 如果值为空且不是必填，跳过其他验证
                if (string.IsNullOrWhiteSpace(value))
                {
                    return result;
                }

                // 格式验证
                if (!string.IsNullOrEmpty(fieldConfig.FieldAttribute.Format))
                {
                    var formatValidation = ValidateFormat(value, fieldConfig.FieldAttribute.Format, fieldConfig.PropertyInfo.PropertyType);
                    if (!formatValidation.IsValid)
                    {
                        result.IsValid = false;
                        result.ErrorMessage = formatValidation.ErrorMessage;
                        result.ErrorType = ExcelImportErrorType.ValidationFailed;
                        return result;
                    }
                }

                // 正则表达式验证
                if (!string.IsNullOrEmpty(fieldConfig.FieldAttribute.ValidationRegex))
                {
                    var regexValidation = ValidateRegex(value, fieldConfig.FieldAttribute.ValidationRegex);
                    if (!regexValidation.IsValid)
                    {
                        result.IsValid = false;
                        result.ErrorMessage = fieldConfig.FieldAttribute.GetEffectiveValidationErrorMessage(fieldConfig.PropertyInfo.Name);
                        result.ErrorType = ExcelImportErrorType.ValidationFailed;
                        return result;
                    }
                }

                // 类型兼容性验证
                var typeValidation = ValidateTypeCompatibility(value, fieldConfig.PropertyInfo.PropertyType);
                if (!typeValidation.IsValid)
                {
                    result.IsValid = false;
                    result.ErrorMessage = typeValidation.ErrorMessage;
                    result.ErrorType = ExcelImportErrorType.DataTypeConversion;
                    return result;
                }

                return result;
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.ErrorMessage = $"验证过程发生错误: {ex.Message}";
                result.ErrorType = ExcelImportErrorType.SystemError;
                return result;
            }
        }

        /// <summary>
        /// 批量验证多个字段
        /// </summary>
        /// <param name="rowData">行数据</param>
        /// <param name="fieldConfigs">字段配置</param>
        /// <param name="rowNumber">行号</param>
        /// <returns>验证结果列表</returns>
        public List<ValidationResult> ValidateRow(Dictionary<string, string> rowData, List<ExcelFieldConfiguration> fieldConfigs, int rowNumber)
        {
            var results = new List<ValidationResult>();

            foreach (var fieldConfig in fieldConfigs)
            {
                var columnName = fieldConfig.EffectiveColumnName;
                var value = rowData.ContainsKey(columnName) ? rowData[columnName] : string.Empty;

                ValidationResult validationResult = null;
                try
                {
                    validationResult = ValidateField(value, fieldConfig, rowNumber);
                }
                catch (Exception ex)
                {
                    // 捕获验证过程中的任何异常，转换为验证失败结果
                    validationResult = new ValidationResult
                    {
                        IsValid = false,
                        FieldName = fieldConfig.PropertyInfo.Name,
                        RowNumber = rowNumber,
                        OriginalValue = value,
                        ErrorMessage = $"验证异常: {ex.Message}",
                        ErrorType = ExcelImportErrorType.SystemError
                    };
                }

                results.Add(validationResult);
            }

            return results;
        }

        #region 私有验证方法

        /// <summary>
        /// 验证格式
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="format">格式字符串</param>
        /// <param name="targetType">目标类型</param>
        /// <returns>验证结果</returns>
        private ValidationResult ValidateFormat(string value, string format, Type targetType)
        {
            var result = new ValidationResult { IsValid = true };

            try
            {
                // 处理可空类型
                var actualType = targetType;
                if (targetType.IsGenericType && targetType.GetGenericTypeDefinition() == typeof(Nullable<>))
                {
                    actualType = Nullable.GetUnderlyingType(targetType);
                }

                // 日期格式验证
                if (actualType == typeof(DateTime))
                {
                    // **修复**: 先尝试日期修正，然后再验证格式
                    string valueToValidate = value;

                    // 尝试修正无效的日期时间值（如秒数=60）
                    if (TryCorrectDateTimeForValidation(value, out string correctedValue))
                    {
                        valueToValidate = correctedValue;
                    }

                    if (!DateTime.TryParseExact(valueToValidate, format, null, System.Globalization.DateTimeStyles.None, out _))
                    {
                        result.IsValid = false;
                        result.ErrorMessage = $"日期格式错误，期望格式: {format}";
                    }
                }
                // 数值格式验证
                else if (actualType == typeof(decimal) || actualType == typeof(double) || actualType == typeof(float))
                {
                    // 简化的数值格式验证
                    if (!decimal.TryParse(value, out _))
                    {
                        result.IsValid = false;
                        result.ErrorMessage = $"数值格式错误，期望格式: {format}";
                    }
                }
                // 枚举格式验证
                else if (actualType.IsEnum)
                {
                    bool isValid = false;
                    
                    // 处理枚举映射格式 "Active=有效;Inactive=无效"
                    if (format.Contains("=") && format.Contains(";"))
                    {
                        var mappings = ParseEnumMappings(format);
                        
                        // 创建反向映射：显示值 -> 枚举名
                        var reverseMappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var mapping in mappings)
                        {
                            reverseMappings[mapping.Value] = mapping.Key;
                        }
                        
                        // 检查显示值或枚举名是否有效
                        isValid = reverseMappings.ContainsKey(value) || mappings.ContainsKey(value);
                    }
                    
                    // 如果映射验证失败，尝试直接枚举验证
                    if (!isValid)
                    {
                        try
                        {
                            // 检查是否为有效的枚举名（忽略大小写）
                            isValid = Enum.IsDefined(actualType, value) || 
                                     Enum.GetNames(actualType).Any(name => name.Equals(value, StringComparison.OrdinalIgnoreCase));
                        }
                        catch
                        {
                            isValid = false;
                        }
                    }
                    
                    if (!isValid)
                    {
                        result.IsValid = false;
                        if (format.Contains("=") && format.Contains(";"))
                        {
                            var mappings = ParseEnumMappings(format);
                            result.ErrorMessage = $"枚举值错误，有效值: {string.Join(", ", mappings.Values)} 或 {string.Join(", ", Enum.GetNames(actualType))}";
                        }
                        else
                        {
                            result.ErrorMessage = $"枚举值错误，有效值: {string.Join(", ", Enum.GetNames(actualType))}";
                        }
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.ErrorMessage = $"格式验证失败: {ex.Message}";
                return result;
            }
        }

        /// <summary>
        /// 验证正则表达式
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="pattern">正则表达式</param>
        /// <returns>验证结果</returns>
        private ValidationResult ValidateRegex(string value, string pattern)
        {
            var result = new ValidationResult { IsValid = true };

            try
            {
                // 从缓存获取或创建正则表达式(线程安全)
                if (!_regexCache.TryGetValue(pattern, out var regex))
                {
                    regex = new Regex(pattern, RegexOptions.Compiled);
                    _regexCache.TryAdd(pattern, regex);
                }

                if (!regex.IsMatch(value))
                {
                    result.IsValid = false;
                    result.ErrorMessage = "格式不符合要求";
                }

                return result;
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.ErrorMessage = $"正则验证失败: {ex.Message}";
                return result;
            }
        }

        /// <summary>
        /// 验证类型兼容性
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="targetType">目标类型</param>
        /// <returns>验证结果</returns>
        private ValidationResult ValidateTypeCompatibility(string value, Type targetType)
        {
            var result = new ValidationResult { IsValid = true };

            try
            {
                // 处理可空类型
                var actualType = targetType;
                if (targetType.IsGenericType && targetType.GetGenericTypeDefinition() == typeof(Nullable<>))
                {
                    actualType = Nullable.GetUnderlyingType(targetType);
                }

                // 字符串类型总是兼容
                if (actualType == typeof(string))
                {
                    return result;
                }

                // 尝试转换以验证兼容性
                if (actualType == typeof(DateTime))
                {
                    // **修复**: 先尝试日期修正，然后再验证类型兼容性
                    string valueToValidate = value;

                    // 尝试修正无效的日期时间值（如秒数=60）
                    if (TryCorrectDateTimeForValidation(value, out string correctedValue))
                    {
                        valueToValidate = correctedValue;
                    }

                    if (!DateTime.TryParse(valueToValidate, out _))
                    {
                        result.IsValid = false;
                        result.ErrorMessage = "无法转换为日期时间格式";
                    }
                }
                else if (actualType == typeof(bool))
                {
                    var lowerValue = value.ToLower().Trim();
                    if (!(lowerValue == "true" || lowerValue == "false" ||
                          lowerValue == "1" || lowerValue == "0" ||
                          lowerValue == "是" || lowerValue == "否" ||
                          lowerValue == "yes" || lowerValue == "no"))
                    {
                        result.IsValid = false;
                        result.ErrorMessage = "无法转换为布尔值，有效值: true/false, 1/0, 是/否, yes/no";
                    }
                }
                else if (actualType.IsEnum)
                {
                    // 枚举类型：跳过类型兼容性验证
                    // 原因：枚举显示值（如"有效"）会在 DataConverter 中通过映射转换
                    // 在这里验证会导致合法的显示值被拒绝
                    // 格式验证（ValidateFormat）已经包含了枚举映射的验证逻辑
                    return result;
                }
                else if (actualType.IsPrimitive || actualType == typeof(decimal))
                {
                    try
                    {
                        Convert.ChangeType(value, actualType);
                    }
                    catch
                    {
                        result.IsValid = false;
                        result.ErrorMessage = $"无法转换为 {actualType.Name} 类型";
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.ErrorMessage = $"类型兼容性验证失败: {ex.Message}";
                return result;
            }
        }

        /// <summary>
        /// 解析枚举映射
        /// </summary>
        /// <param name="format">格式字符串</param>
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

        /// <summary>
        /// 尝试修正无效的日期时间值（用于验证阶段）
        /// 这是 DataConverter.TryCorrectDateTime 的副本，用于在验证阶段也能应用日期修正
        /// </summary>
        private bool TryCorrectDateTimeForValidation(string value, out string correctedValue)
        {
            correctedValue = value;

            try
            {
                // 匹配日期时间格式：yyyy-MM-dd HH:mm:ss 或类似格式
                // 支持的分隔符：- / . 和空格
                var match = System.Text.RegularExpressions.Regex.Match(value,
                    @"^(\d{4})[-/\.](\d{1,2})[-/\.](\d{1,2})[\sT](\d{1,2}):(\d{1,2}):(\d{1,2})");

                if (match.Success)
                {
                    int year = int.Parse(match.Groups[1].Value);
                    int month = int.Parse(match.Groups[2].Value);
                    int day = int.Parse(match.Groups[3].Value);
                    int hour = int.Parse(match.Groups[4].Value);
                    int minute = int.Parse(match.Groups[5].Value);
                    int second = int.Parse(match.Groups[6].Value);

                    bool needsCorrection = false;

                    // 修正无效的秒数（>= 60）
                    if (second >= 60)
                    {
                        minute += second / 60;
                        second = second % 60;
                        needsCorrection = true;
                    }

                    // 修正无效的分钟数（>= 60）
                    if (minute >= 60)
                    {
                        hour += minute / 60;
                        minute = minute % 60;
                        needsCorrection = true;
                    }

                    // 修正无效的小时数（>= 24）
                    if (hour >= 24)
                    {
                        day += hour / 24;
                        hour = hour % 24;
                        needsCorrection = true;
                    }

                    // 修正无效的天数（根据月份）
                    if (day > DateTime.DaysInMonth(year, month))
                    {
                        day = DateTime.DaysInMonth(year, month);
                        needsCorrection = true;
                    }

                    // 如果需要修正，重新构建日期时间字符串
                    if (needsCorrection)
                    {
                        // 保持原始格式
                        correctedValue = $"{year:D4}-{month:D2}-{day:D2} {hour:D2}:{minute:D2}:{second:D2}";
                        System.Diagnostics.Debug.WriteLine($"[DataValidator] 验证阶段修正日期时间: {value} -> {correctedValue}");
                        return true;
                    }
                }

                return false;
            }
            catch
            {
                // 如果解析失败，返回原始值
                return false;
            }
        }
    }

    /// <summary>
    /// 验证结果
    /// </summary>
    public class ValidationResult
    {
        /// <summary>
        /// 是否验证通过
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// 字段名
        /// </summary>
        public string FieldName { get; set; }

        /// <summary>
        /// 行号
        /// </summary>
        public int RowNumber { get; set; }

        /// <summary>
        /// 原始值
        /// </summary>
        public string OriginalValue { get; set; }

        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 错误类型
        /// </summary>
        public ExcelImportErrorType ErrorType { get; set; }

        /// <summary>
        /// 转换为ExcelImportError
        /// </summary>
        /// <returns>Excel导入错误</returns>
        public ExcelImportError ToExcelImportError()
        {
            return new ExcelImportError
            {
                RowNumber = RowNumber,
                FieldName = FieldName,
                ErrorMessage = ErrorMessage,
                OriginalValue = OriginalValue,
                ErrorType = ErrorType
            };
        }
    }
}