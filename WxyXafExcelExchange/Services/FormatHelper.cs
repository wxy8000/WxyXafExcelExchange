using System;
using System.Globalization;
using System.Reflection;
using Wxy.Xaf.ExcelExchange;

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// 格式解析结果
    /// </summary>
    public class FormatParseResult
    {
        /// <summary>
        /// 是否解析成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 解析后的值
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 创建成功结果
        /// </summary>
        /// <param name="value">解析后的值</param>
        /// <returns>成功结果</returns>
        public static FormatParseResult CreateSuccess(object value)
        {
            return new FormatParseResult { Success = true, Value = value };
        }

        /// <summary>
        /// 创建失败结果
        /// </summary>
        /// <param name="errorMessage">错误消息</param>
        /// <returns>失败结果</returns>
        public static FormatParseResult CreateFailure(string errorMessage)
        {
            return new FormatParseResult { Success = false, ErrorMessage = errorMessage };
        }
    }

    /// <summary>
    /// 自定义格式解析辅助类
    /// </summary>
    public static class FormatHelper
    {
        /// <summary>
        /// 使用自定义格式解析值
        /// </summary>
        /// <param name="value">要解析的字符串值</param>
        /// <param name="targetType">目标类型</param>
        /// <param name="customFormat">自定义格式字符串</param>
        /// <returns>解析结果</returns>
        public static FormatParseResult ParseWithCustomFormat(string value, Type targetType, string customFormat)
        {
            if (string.IsNullOrWhiteSpace(value))
                return FormatParseResult.CreateSuccess(GetDefaultValue(targetType));

            // 处理可空类型
            var underlyingType = Nullable.GetUnderlyingType(targetType) ?? targetType;

            try
            {
                // 日期类型的自定义格式解析 - 增强版本
                if (underlyingType == typeof(DateTime))
                {
                    // 首先尝试自定义格式
                    if (!string.IsNullOrEmpty(customFormat))
                    {
                        if (DateTime.TryParseExact(value, customFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateResult))
                        {
                            return FormatParseResult.CreateSuccess(dateResult);
                        }
                    }
                    
                    // 尝试多种常见的日期格式
                    var commonDateFormats = new[]
                    {
                        "yyyy-MM-dd",           // ISO 8601
                        "yyyy/MM/dd",           // 年/月/日
                        "dd/MM/yyyy",           // 日/月/年
                        "MM/dd/yyyy",           // 月/日/年
                        "dd-MM-yyyy",           // 日-月-年
                        "MM-dd-yyyy",           // 月-日-年
                        "yyyy.MM.dd",           // 年.月.日
                        "dd.MM.yyyy",           // 日.月.年
                        "yyyy年MM月dd日",        // 中文格式
                        "yyyy-MM-dd HH:mm:ss",  // 带时间的ISO格式
                        "yyyy/MM/dd HH:mm:ss",  // 带时间的斜杠格式
                        "dd/MM/yyyy HH:mm:ss",  // 欧洲格式带时间
                        "MM/dd/yyyy HH:mm:ss",  // 美国格式带时间
                        "yyyy-MM-dd HH:mm",     // 不带秒的时间
                        "yyyy/MM/dd HH:mm",     // 不带秒的斜杠格式
                        "yyyyMMdd",             // 紧凑格式
                        "ddMMyyyy",             // 紧凑欧洲格式
                        "MMddyyyy"              // 紧凑美国格式
                    };

                    foreach (var format in commonDateFormats)
                    {
                        if (DateTime.TryParseExact(value, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime formatResult))
                        {
                            return FormatParseResult.CreateSuccess(formatResult);
                        }
                    }
                    
                    // 如果所有格式都失败，尝试标准解析
                    if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime standardDateResult))
                    {
                        return FormatParseResult.CreateSuccess(standardDateResult);
                    }
                    
                    // 最后尝试当前文化的解析
                    if (DateTime.TryParse(value, out DateTime currentCultureResult))
                    {
                        return FormatParseResult.CreateSuccess(currentCultureResult);
                    }
                    
                    return FormatParseResult.CreateFailure($"无法将 '{value}' 解析为日期格式");
                }

                // 数字类型的自定义格式解析
                if (IsNumericType(underlyingType))
                {
                    var numericResult = ParseNumericWithFormat(value, underlyingType, customFormat);
                    return FormatParseResult.CreateSuccess(numericResult);
                }

                // 其他类型使用标准转换
                var result = Convert.ChangeType(value, underlyingType, CultureInfo.InvariantCulture);
                return FormatParseResult.CreateSuccess(result);
            }
            catch (Exception ex)
            {
                return FormatParseResult.CreateFailure($"无法将 '{value}' 转换为类型 '{targetType.Name}': {ex.Message}");
            }
        }

        /// <summary>
        /// 检查类型是否为数字类型
        /// </summary>
        /// <param name="type">要检查的类型</param>
        /// <returns>是否为数字类型</returns>
        private static bool IsNumericType(Type type)
        {
            return type == typeof(int) || type == typeof(long) || type == typeof(short) || type == typeof(byte) ||
                   type == typeof(uint) || type == typeof(ulong) || type == typeof(ushort) || type == typeof(sbyte) ||
                   type == typeof(decimal) || type == typeof(double) || type == typeof(float);
        }

        /// <summary>
        /// 使用自定义格式解析数字
        /// </summary>
        /// <param name="value">要解析的字符串值</param>
        /// <param name="targetType">目标数字类型</param>
        /// <param name="customFormat">自定义格式字符串</param>
        /// <returns>解析后的数字值</returns>
        private static object ParseNumericWithFormat(string value, Type targetType, string customFormat)
        {
            // 移除可能的格式化字符（如千位分隔符、货币符号等）
            var cleanValue = CleanNumericValue(value);

            if (targetType == typeof(decimal))
            {
                return decimal.Parse(cleanValue, NumberStyles.Any, CultureInfo.InvariantCulture);
            }
            if (targetType == typeof(double))
            {
                return double.Parse(cleanValue, NumberStyles.Any, CultureInfo.InvariantCulture);
            }
            if (targetType == typeof(float))
            {
                return float.Parse(cleanValue, NumberStyles.Any, CultureInfo.InvariantCulture);
            }
            if (targetType == typeof(int))
            {
                return int.Parse(cleanValue, NumberStyles.Integer, CultureInfo.InvariantCulture);
            }
            if (targetType == typeof(long))
            {
                return long.Parse(cleanValue, NumberStyles.Integer, CultureInfo.InvariantCulture);
            }
            if (targetType == typeof(short))
            {
                return short.Parse(cleanValue, NumberStyles.Integer, CultureInfo.InvariantCulture);
            }
            if (targetType == typeof(byte))
            {
                return byte.Parse(cleanValue, NumberStyles.Integer, CultureInfo.InvariantCulture);
            }

            // 其他数字类型使用标准转换
            return Convert.ChangeType(cleanValue, targetType, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// 清理数字值中的格式化字符 - 改进版本，支持更多货币符号和千分位分隔符
        /// </summary>
        /// <param name="value">原始数字字符串</param>
        /// <returns>清理后的数字字符串</returns>
        public static string CleanNumericValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return value;

            // 移除常见的格式化字符
            var cleanValue = value.Trim()
                       .Replace(",", "")    // 千位分隔符（英文）
                       .Replace(".", ".")   // 保留小数点
                       .Replace(" ", "")    // 空格分隔符（法语等）
                       .Replace("'", "")    // 撇号分隔符（瑞士等）
                       .Replace("_", "")    // 下划线分隔符
                       .Replace("$", "")    // 美元符号
                       .Replace("¥", "")    // 人民币符号
                       .Replace("€", "")    // 欧元符号
                       .Replace("£", "")    // 英镑符号
                       .Replace("₹", "")    // 印度卢比符号
                       .Replace("₽", "")    // 俄罗斯卢布符号
                       .Replace("₩", "")    // 韩元符号
                       .Replace("¢", "")    // 美分符号
                       .Replace("₡", "")    // 哥伦比亚比索符号
                       .Replace("₦", "")    // 尼日利亚奈拉符号
                       .Replace("₨", "")    // 巴基斯坦卢比符号
                       .Replace("₪", "")    // 以色列新谢克尔符号
                       .Replace("₫", "")    // 越南盾符号
                       .Replace("₱", "")    // 菲律宾比索符号
                       .Replace("₲", "")    // 巴拉圭瓜拉尼符号
                       .Replace("₴", "")    // 乌克兰格里夫纳符号
                       .Replace("₵", "")    // 加纳塞地符号
                       .Replace("₶", "")    // 法国法郎符号
                       .Replace("₷", "")    // 西班牙比塞塔符号
                       .Replace("₸", "")    // 哈萨克斯坦坚戈符号
                       .Replace("₹", "")    // 印度卢比符号
                       .Replace("₺", "")    // 土耳其里拉符号
                       .Replace("₻", "")    // 北欧马克符号
                       .Replace("₼", "")    // 阿塞拜疆马纳特符号
                       .Replace("₽", "")    // 俄罗斯卢布符号
                       .Replace("₾", "")    // 格鲁吉亚拉里符号
                       .Replace("₿", "")    // 比特币符号
                       .Replace("%", "");   // 百分号

            // 处理负号位置（可能在前面或后面）
            bool isNegative = cleanValue.StartsWith("-") || cleanValue.EndsWith("-") || 
                             cleanValue.StartsWith("(") && cleanValue.EndsWith(")");
            
            if (isNegative)
            {
                cleanValue = cleanValue.Replace("-", "").Replace("(", "").Replace(")", "").Trim();
                cleanValue = "-" + cleanValue;
            }

            return cleanValue;
        }

        /// <summary>
        /// 获取类型的默认值
        /// </summary>
        /// <param name="type">类型</param>
        /// <returns>默认值</returns>
        private static object GetDefaultValue(Type type)
        {
            if (type.IsValueType)
            {
                return Activator.CreateInstance(type);
            }
            return null;
        }

        /// <summary>
        /// 使用自定义格式格式化值（用于导出）
        /// </summary>
        /// <param name="value">要格式化的值</param>
        /// <param name="customFormat">自定义格式字符串</param>
        /// <returns>格式化后的字符串</returns>
        public static string FormatWithCustomFormat(object value, string customFormat)
        {
            if (value == null)
                return string.Empty;

            if (string.IsNullOrEmpty(customFormat))
                return value.ToString();

            try
            {
                // 日期格式化
                if (value is DateTime dateValue)
                {
                    return dateValue.ToString(customFormat, CultureInfo.InvariantCulture);
                }

                // 数字格式化
                if (IsNumericType(value.GetType()))
                {
                    if (value is decimal decimalValue)
                        return decimalValue.ToString(customFormat, CultureInfo.InvariantCulture);
                    if (value is double doubleValue)
                        return doubleValue.ToString(customFormat, CultureInfo.InvariantCulture);
                    if (value is float floatValue)
                        return floatValue.ToString(customFormat, CultureInfo.InvariantCulture);
                    if (value is int intValue)
                        return intValue.ToString(customFormat, CultureInfo.InvariantCulture);
                    if (value is long longValue)
                        return longValue.ToString(customFormat, CultureInfo.InvariantCulture);
                }

                // 其他类型使用标准格式化
                return value.ToString();
            }
            catch (Exception)
            {
                // 如果自定义格式化失败，返回标准字符串表示
                return value.ToString();
            }
        }

        /// <summary>
        /// 验证自定义格式字符串是否有效
        /// </summary>
        /// <param name="format">格式字符串</param>
        /// <param name="targetType">目标类型</param>
        /// <returns>是否有效</returns>
        public static bool IsValidFormat(string format, Type targetType)
        {
            if (string.IsNullOrEmpty(format))
                return true;

            var underlyingType = Nullable.GetUnderlyingType(targetType) ?? targetType;

            try
            {
                if (underlyingType == typeof(DateTime))
                {
                    // 测试日期格式
                    DateTime.Now.ToString(format, CultureInfo.InvariantCulture);
                    return true;
                }

                if (IsNumericType(underlyingType))
                {
                    // 测试数字格式
                    if (underlyingType == typeof(decimal))
                        ((decimal)123.45).ToString(format, CultureInfo.InvariantCulture);
                    else if (underlyingType == typeof(double))
                        ((double)123.45).ToString(format, CultureInfo.InvariantCulture);
                    else if (underlyingType == typeof(int))
                        ((int)123).ToString(format, CultureInfo.InvariantCulture);
                    
                    return true;
                }

                return true; // 其他类型不验证格式
            }
            catch
            {
                return false;
            }
        }
    }
}