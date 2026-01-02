using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Wxy.Xaf.ExcelExchange.Configuration;
using Wxy.Xaf.ExcelExchange.Services;

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// 数据转换服务，负责在Excel数据和对象属性之间进行转换
    /// </summary>
    public class DataConverter
    {
        private readonly ConcurrentDictionary<string, MethodInfo> _converterMethodCache = new ConcurrentDictionary<string, MethodInfo>();

        /// <summary>
        /// 将Excel字符串值转换为对象属性值
        /// </summary>
        /// <param name="value">Excel字符串值</param>
        /// <param name="fieldConfig">字段配置</param>
        /// <returns>转换结果</returns>
        public ConvertResult ConvertFromExcel(string value, ExcelFieldConfiguration fieldConfig)
        {
            var result = new ConvertResult
            {
                IsSuccess = true,
                OriginalValue = value
            };

            try
            {
                var targetType = fieldConfig.PropertyInfo.PropertyType;

                // 处理空值
                if (string.IsNullOrEmpty(value))
                {
                    return HandleNullValue(fieldConfig, result);
                }

                // 检查是否有自定义导入转换方法
                if (!string.IsNullOrEmpty(fieldConfig.FieldAttribute.ImportConvertMethod))
                {
                    return ConvertUsingCustomMethod(value, fieldConfig.FieldAttribute.ImportConvertMethod, result);
                }

                // 使用默认转换逻辑
                result.ConvertedValue = ConvertToType(value, targetType, fieldConfig);
                return result;
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.ErrorMessage = $"转换失败: {ex.Message}";
                return result;
            }
        }

        /// <summary>
        /// 将对象属性值转换为Excel显示值
        /// </summary>
        /// <param name="value">对象属性值</param>
        /// <param name="fieldConfig">字段配置</param>
        /// <returns>转换结果</returns>
        public ConvertResult ConvertToExcel(object value, ExcelFieldConfiguration fieldConfig)
        {
            var result = new ConvertResult
            {
                IsSuccess = true,
                OriginalValue = value
            };

            try
            {
                // 处理空值
                if (value == null)
                {
                    result.ConvertedValue = fieldConfig.FieldAttribute.ExportNullValueDisplay ?? "";
                    return result;
                }

                // 检查是否有自定义导出转换方法
                if (!string.IsNullOrEmpty(fieldConfig.FieldAttribute.ExportConvertMethod))
                {
                    return ConvertUsingCustomMethod(value, fieldConfig.FieldAttribute.ExportConvertMethod, result);
                }

                // 使用默认转换逻辑
                result.ConvertedValue = ConvertFromType(value, fieldConfig);
                return result;
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.ErrorMessage = $"转换失败: {ex.Message}";
                return result;
            }
        }

        #region 私有转换方法

        /// <summary>
        /// 处理空值
        /// </summary>
        /// <param name="fieldConfig">字段配置</param>
        /// <param name="result">转换结果</param>
        /// <returns>转换结果</returns>
        private ConvertResult HandleNullValue(ExcelFieldConfiguration fieldConfig, ConvertResult result)
        {
            var targetType = fieldConfig.PropertyInfo.PropertyType;

            // 检查是否有默认值替换
            if (fieldConfig.FieldAttribute.ImportNullValueReplacement != null)
            {
                try
                {
                    var replacement = fieldConfig.FieldAttribute.ImportNullValueReplacement;
                    
                    // 如果替换值是Type类型，表示需要创建默认实例
                    if (replacement is Type typeReplacement)
                    {
                        if (typeReplacement == typeof(DateTime))
                        {
                            result.ConvertedValue = DateTime.Now;
                        }
                        else
                        {
                            result.ConvertedValue = Activator.CreateInstance(typeReplacement);
                        }
                    }
                    else
                    {
                        result.ConvertedValue = ConvertToType(replacement.ToString(), targetType, fieldConfig);
                    }

                    result.HasWarning = true;
                    result.WarningMessage = "使用了默认值替换";
                    return result;
                }
                catch (Exception ex)
                {
                    result.IsSuccess = false;
                    result.ErrorMessage = $"默认值替换失败: {ex.Message}";
                    return result;
                }
            }

            // 处理可空类型
            if (targetType.IsGenericType && targetType.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                result.ConvertedValue = null;
                return result;
            }

            // 字符串类型
            if (targetType == typeof(string))
            {
                result.ConvertedValue = string.Empty;
                return result;
            }

            // 值类型使用默认值
            if (targetType.IsValueType)
            {
                result.ConvertedValue = Activator.CreateInstance(targetType);
                result.HasWarning = true;
                result.WarningMessage = "空值已转换为默认值";
                return result;
            }

            // 引用类型设为null
            result.ConvertedValue = null;
            return result;
        }

        /// <summary>
        /// 使用自定义方法转换
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="methodName">方法名</param>
        /// <param name="result">转换结果</param>
        /// <returns>转换结果</returns>
        private ConvertResult ConvertUsingCustomMethod(object value, string methodName, ConvertResult result)
        {
            try
            {
                // 从缓存获取或查找方法(线程安全)
                if (!_converterMethodCache.TryGetValue(methodName, out var method))
                {
                    method = FindConverterMethod(methodName);
                    if (method != null)
                    {
                        _converterMethodCache.TryAdd(methodName, method);
                    }
                }

                if (method == null)
                {
                    result.IsSuccess = false;
                    result.ErrorMessage = $"未找到转换方法: {methodName}";
                    return result;
                }

                // 调用转换方法
                result.ConvertedValue = method.Invoke(null, new[] { value });
                return result;
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.ErrorMessage = $"自定义转换方法调用失败: {ex.Message}";
                return result;
            }
        }

        /// <summary>
        /// 查找转换方法
        /// </summary>
        /// <param name="methodName">方法名</param>
        /// <returns>方法信息</returns>
        private MethodInfo FindConverterMethod(string methodName)
        {
            try
            {
                // 在当前程序集中查找静态方法
                var assemblies = new[] { Assembly.GetExecutingAssembly(), Assembly.GetCallingAssembly() };
                
                foreach (var assembly in assemblies)
                {
                    foreach (var type in assembly.GetTypes())
                    {
                        var method = type.GetMethod(methodName, BindingFlags.Public | BindingFlags.Static);
                        if (method != null)
                        {
                            return method;
                        }
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 转换为指定类型
        /// </summary>
        /// <param name="value">字符串值</param>
        /// <param name="targetType">目标类型</param>
        /// <param name="fieldConfig">字段配置</param>
        /// <returns>转换后的值</returns>
        private object ConvertToType(string value, Type targetType, ExcelFieldConfiguration fieldConfig)
        {
            // 处理可空类型
            Type actualType = targetType;
            if (targetType.IsGenericType && targetType.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                actualType = Nullable.GetUnderlyingType(targetType);
            }

            // 字符串类型直接返回
            if (actualType == typeof(string))
            {
                return value;
            }

            // DataDictionaryItem 类型处理 - 返回字符串值，后续由 ObjectSpace 查找
            if (actualType.Name == "DataDictionaryItem" || actualType.FullName?.Contains("DataDictionaryItem") == true)
            {
                
                // 空值处理
                if (string.IsNullOrWhiteSpace(value))
                {
                    return null;
                }

                // 直接返回字符串值，让后续的属性设置逻辑处理对象查找
                return value;
            }

            // 日期时间类型
            if (actualType == typeof(DateTime))
            {
                // 空值处理
                if (string.IsNullOrWhiteSpace(value))
                {
                    return null;
                }

                // **修复**: 检查是否为 Excel 的无效日期表示（0001-01-01 或类似）
                if (value.Contains("0001-01-01") || value.Contains("1/1/0001") || value.Contains("0001/1/1"))
                {
                    return null;
                }

                // **新增**: 尝试修正无效的日期时间值（如秒数=60）
                DateTime dateValue;
                string correctedValue = value;

                // 检测并修正无效的秒数（60, 61等）
                if (TryCorrectDateTime(value, out correctedValue))
                {
                    System.Diagnostics.Debug.WriteLine($"[DataConverter] 修正日期时间: {value} -> {correctedValue}");
                }

                // 尝试多种日期格式
                var dateFormats = new[]
                {
                    fieldConfig.FieldAttribute.Format,
                    // 日期时间格式（带时间部分）
                    "yyyy-MM-dd HH:mm:ss",
                    "yyyy/MM/dd HH:mm:ss",
                    "yyyy/M/d HH:mm:ss",
                    "yyyy-M-d HH:mm:ss",
                    // 纯日期格式
                    "yyyy/M/d",
                    "yyyy-M-d",
                    "yyyy/MM/dd",
                    "yyyy-MM-dd",
                    "M/d/yyyy",
                    "MM/dd/yyyy",
                    "d/M/yyyy",
                    "dd/MM/yyyy"
                }.Where(f => !string.IsNullOrEmpty(f)).ToArray();

                // 首先尝试指定格式（使用修正后的值）
                foreach (var format in dateFormats)
                {
                    if (DateTime.TryParseExact(correctedValue, format, null, System.Globalization.DateTimeStyles.None, out dateValue))
                    {
                        // **修复**: 检查是否为 DateTime.MinValue，SQL Server 不支持 0001-01-01
                        if (dateValue == DateTime.MinValue || dateValue.Year < 1753)
                        {
                            return null;
                        }
                        return dateValue;
                    }
                }

                // 然后尝试通用解析（使用修正后的值）
                if (DateTime.TryParse(correctedValue, out dateValue))
                {
                    // **修复**: 检查是否为 DateTime.MinValue 或年份小于1753
                    if (dateValue == DateTime.MinValue || dateValue.Year < 1753)
                    {
                        return null;
                    }
                    return dateValue;
                }

                throw new FormatException($"无法将 '{value}' 转换为日期时间格式，期望格式: {fieldConfig.FieldAttribute.Format ?? "yyyy/M/d"}");
            }

            // 布尔类型
            if (actualType == typeof(bool))
            {
                var lowerValue = value.ToLower().Trim();
                if (lowerValue == "true" || lowerValue == "1" || lowerValue == "是" || lowerValue == "yes")
                {
                    return true;
                }
                if (lowerValue == "false" || lowerValue == "0" || lowerValue == "否" || lowerValue == "no")
                {
                    return false;
                }
                throw new FormatException($"无法将 '{value}' 转换为布尔值");
            }

            // 枚举类型
            if (actualType.IsEnum)
            {
                // 对值进行 trim，去除前后空格
                var trimmedValue = value.Trim();
                
                // 检查是否有枚举映射格式
                if (!string.IsNullOrEmpty(fieldConfig.FieldAttribute.Format) &&
                    fieldConfig.FieldAttribute.Format.Contains("="))
                {
                    var reverseMappings = ParseEnumMappings(fieldConfig.FieldAttribute.Format);
                    var mappings = reverseMappings.ToDictionary(kv => kv.Value, kv => kv.Key);

                    // 尝试通过显示值查找枚举名（使用 trim 后的值）
                    if (reverseMappings.TryGetValue(trimmedValue, out string enumName))
                    {
                        return Enum.Parse(actualType, enumName, true);
                    }
                    else
                    {
                        // 尝试直接使用显示值作为枚举名
                        if (Enum.IsDefined(actualType, trimmedValue))
                        {
                            return Enum.Parse(actualType, trimmedValue, true);
                        }
                    }

                }

                // 尝试直接解析枚举（忽略大小写）
                try
                {
                    return Enum.Parse(actualType, trimmedValue, true);
                }
                catch
                {
                    throw new FormatException($"无法将 '{trimmedValue}' 转换为枚举类型 {actualType.Name}，可用值: {string.Join(", ", Enum.GetNames(actualType))}");
                }
            }

            // 其他类型使用通用转换
            return Convert.ChangeType(value, actualType);
        }

        /// <summary>
        /// 从类型转换为字符串
        /// </summary>
        /// <param name="value">对象值</param>
        /// <param name="fieldConfig">字段配置</param>
        /// <returns>字符串值</returns>
        private string ConvertFromType(object value, ExcelFieldConfiguration fieldConfig)
        {
            if (value == null)
            {
                return fieldConfig.FieldAttribute.ExportNullValueDisplay ?? "";
            }

            var valueType = value.GetType();

            // 日期时间格式化
            if (value is DateTime dateTime)
            {
                var format = fieldConfig.FieldAttribute.Format ?? "yyyy-MM-dd HH:mm:ss";
                return dateTime.ToString(format);
            }

            // 数值格式化
            if (value is decimal || value is double || value is float)
            {
                var format = fieldConfig.FieldAttribute.Format ?? "0.##";
                return ((IFormattable)value).ToString(format, null);
            }

            // 布尔值转换
            if (value is bool boolValue)
            {
                // 检查是否有自定义格式
                if (!string.IsNullOrEmpty(fieldConfig.FieldAttribute.Format))
                {
                    var mappings = ParseBooleanMappings(fieldConfig.FieldAttribute.Format);
                    return boolValue ? mappings.TrueValue : mappings.FalseValue;
                }
                return boolValue ? "是" : "否";
            }

            // 枚举转换
            if (valueType.IsEnum)
            {
                // 检查是否有枚举映射格式
                if (!string.IsNullOrEmpty(fieldConfig.FieldAttribute.Format) &&
                    fieldConfig.FieldAttribute.Format.Contains("="))
                {
                    var mappings = ParseEnumMappings(fieldConfig.FieldAttribute.Format);
                    var enumName = value.ToString();
                    if (mappings.ContainsKey(enumName))
                    {
                        return mappings[enumName];
                    }
                }
                return value.ToString();
            }

            // XPO 集合类型处理（XPCollection）- 展开为文本
            if (valueType.IsGenericType && valueType.GetGenericTypeDefinition().Name.Contains("XPCollection"))
            {
                return ConvertXPCollectionToText(value, fieldConfig);
            }

            // 再次检查非泛型的 XPCollection
            if (valueType.FullName != null && valueType.FullName.Contains("XPCollection"))
            {
                return ConvertXPCollectionToText(value, fieldConfig);
            }

            // XPO 关联对象处理 - 导出代码字段值
            if (valueType.GetInterface("DevExpress.Xpo.IXPObject") != null)
            {
                try
                {
                    var matchFieldName = fieldConfig.FieldAttribute.ReferenceMatchField;

                    // **特殊处理**: DataDictionaryItem 类型自动使用 Name 属性
                    // 如果 ReferenceMatchField 是默认值 "Oid" 或者未设置，则使用 "Name" 作为默认
                    System.Diagnostics.Debug.WriteLine($"[Export] 检查类型: {valueType.Name}, FullName: {valueType.FullName}, ReferenceMatchField: {matchFieldName}");
                    if (valueType.Name == "DataDictionaryItem" || valueType.FullName?.Contains("DataDictionaryItem") == true)
                    {
                        // 对于 DataDictionaryItem，如果没有显式指定非 "Oid" 的 ReferenceMatchField，默认使用 "Name"
                        var propertyName = (string.IsNullOrEmpty(matchFieldName) || matchFieldName == "Oid") ? "Name" : matchFieldName;
                        System.Diagnostics.Debug.WriteLine($"[Export] DataDictionaryItem 使用属性: {propertyName}");

                        var nameProperty = valueType.GetProperty(propertyName);
                        if (nameProperty != null)
                        {
                            var nameValue = nameProperty.GetValue(value);
                            System.Diagnostics.Debug.WriteLine($"[Export] DataDictionaryItem.{propertyName} = {nameValue}");
                            return nameValue?.ToString() ?? "";
                        }
                    }

                    // 如果配置了 ReferenceMatchField，使用指定属性
                    if (!string.IsNullOrEmpty(matchFieldName))
                    {
                        var matchProperty = valueType.GetProperty(matchFieldName);
                        if (matchProperty != null)
                        {
                            var matchValue = matchProperty.GetValue(value);
                            if (matchValue != null)
                            {
                                return matchValue.ToString();
                            }
                        }
                    }

                    // 尝试使用 ToString()
                    var toStringValue = value.ToString();

                    // 如果 ToString() 返回的是类型名称或 GUID，返回空
                    if (toStringValue.Contains(valueType.Name) ||
                        toStringValue.Contains("DevExpress.Xpo") ||
                        Guid.TryParse(toStringValue, out _))
                    {
                        return "";
                    }
                    return toStringValue;
                }
                catch
                {
                    return "";
                }
            }

            // 默认转换
            return value.ToString();
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

        /// <summary>
        /// 解析布尔值映射
        /// </summary>
        /// <param name="format">格式字符串</param>
        /// <returns>布尔值映射</returns>
        private (string TrueValue, string FalseValue) ParseBooleanMappings(string format)
        {
            try
            {
                if (format.Contains("="))
                {
                    var parts = format.Split(';');
                    var trueValue = "是";
                    var falseValue = "否";

                    foreach (var part in parts)
                    {
                        var mapping = part.Split('=');
                        if (mapping.Length == 2)
                        {
                            var key = mapping[0].Trim().ToLower();
                            var value = mapping[1].Trim();
                            
                            if (key == "true" || key == "1")
                            {
                                trueValue = value;
                            }
                            else if (key == "false" || key == "0")
                            {
                                falseValue = value;
                            }
                        }
                    }

                    return (trueValue, falseValue);
                }
            }
            catch
            {
                // 忽略解析错误
            }

            return ("是", "否");
        }

        /// <summary>
        /// 将 XPCollection 转换为文本表示
        /// </summary>
        private string ConvertXPCollectionToText(object collection, ExcelFieldConfiguration fieldConfig)
        {
            try
            {
                if (collection == null)
                    return "";

                // 获取集合中的元素
                var collectionType = collection.GetType();
                var countProperty = collectionType.GetProperty("Count");
                if (countProperty == null)
                    return "";

                int count = (int)countProperty.GetValue(collection);
                if (count == 0)
                    return "无明细";

                // 获取导出格式配置
                var format = fieldConfig.FieldAttribute.CollectionExportFormat ?? "Summary";
                var delimiter = fieldConfig.FieldAttribute.CollectionDelimiter ?? "; ";

                // Summary 模式：汇总为单个文本
                if (format.Equals("Summary", StringComparison.OrdinalIgnoreCase))
                {
                    var items = new List<string>();
                    var getItemMethod = collectionType.GetMethod("GetItemByIndex");

                    for (int i = 0; i < count; i++)
                    {
                        var item = getItemMethod?.Invoke(collection, new object[] { i });
                        if (item != null)
                        {
                            items.Add(FormatCollectionItem(item, fieldConfig));
                        }
                    }

                    return string.Join(delimiter, items);
                }

                // Count 模式：仅显示数量
                if (format.Equals("Count", StringComparison.OrdinalIgnoreCase))
                {
                    return $"{count}条明细";
                }

                // MultiRow 模式：暂不支持（需要在导出层面展开为多行）
                // 返回 Summary 格式作为后备
                var summaryItems = new List<string>();
                var getMethod = collectionType.GetMethod("GetItemByIndex");
                for (int i = 0; i < count; i++)
                {
                    var item = getMethod?.Invoke(collection, new object[] { i });
                    if (item != null)
                    {
                        summaryItems.Add(FormatCollectionItem(item, fieldConfig));
                    }
                }
                return string.Join(delimiter, summaryItems);
            }
            catch
            {
                                return "";
            }
        }

        /// <summary>
        /// 格式化集合中的单个项目
        /// </summary>
        private string FormatCollectionItem(object item, ExcelFieldConfiguration fieldConfig)
        {
            try
            {
                var itemType = item.GetType();
                var displayProps = fieldConfig.FieldAttribute.CollectionDisplayProperties;

                // 如果配置了显示属性
                if (!string.IsNullOrEmpty(displayProps))
                {
                    var propNames = displayProps.Split(',');
                    var parts = new List<string>();

                    foreach (var propName in propNames)
                    {
                        var prop = itemType.GetProperty(propName.Trim());
                        if (prop != null)
                        {
                            var value = prop.GetValue(item);
                            if (value != null)
                            {
                                parts.Add(value.ToString());
                            }
                        }
                    }

                    if (parts.Count > 0)
                    {
                        return string.Join(" ", parts);
                    }
                }

                // 自动检测：尝试找到第一个字符串属性（通常是 Name 或 Title）
                var stringProp = itemType.GetProperties()
                    .FirstOrDefault(p => p.PropertyType == typeof(string) &&
                                       p.CanRead &&
                                       !p.Name.Equals("ClassName"));

                if (stringProp != null)
                {
                    return stringProp.GetValue(item)?.ToString() ?? item.ToString();
                }

                return item.ToString();
            }
            catch
            {
                return item?.ToString() ?? "";
            }
        }

        #endregion

        /// <summary>
        /// 尝试修正无效的日期时间值（如秒数=60）
        /// </summary>
        /// <param name="value">原始日期时间字符串</param>
        /// <param name="correctedValue">修正后的日期时间字符串</param>
        /// <returns>是否进行了修正</returns>
        private bool TryCorrectDateTime(string value, out string correctedValue)
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
    /// 转换结果
    /// </summary>
    public class ConvertResult
    {
        /// <summary>
        /// 是否转换成功
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 原始值
        /// </summary>
        public object OriginalValue { get; set; }

        /// <summary>
        /// 转换后的值
        /// </summary>
        public object ConvertedValue { get; set; }

        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 是否有警告
        /// </summary>
        public bool HasWarning { get; set; }

        /// <summary>
        /// 警告消息
        /// </summary>
        public string WarningMessage { get; set; }

        /// <summary>
        /// 转换为ExcelImportError
        /// </summary>
        /// <param name="rowNumber">行号</param>
        /// <param name="fieldName">字段名</param>
        /// <returns>Excel导入错误</returns>
        public ExcelImportError ToExcelImportError(int rowNumber, string fieldName)
        {
            return new ExcelImportError
            {
                RowNumber = rowNumber,
                FieldName = fieldName,
                ErrorMessage = ErrorMessage,
                OriginalValue = OriginalValue?.ToString(),
                ErrorType = ExcelImportErrorType.DataTypeConversion
            };
        }

        /// <summary>
        /// 转换为ExcelImportWarning
        /// </summary>
        /// <param name="rowNumber">行号</param>
        /// <param name="fieldName">字段名</param>
        /// <returns>Excel导入警告</returns>
        public ExcelImportWarning ToExcelImportWarning(int rowNumber, string fieldName)
        {
            return new ExcelImportWarning
            {
                RowNumber = rowNumber,
                FieldName = fieldName,
                WarningMessage = WarningMessage,
                OriginalValue = OriginalValue?.ToString(),
                ConvertedValue = ConvertedValue?.ToString()
            };
        }
    }
}