using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
#if NET8_0_OR_GREATER
using System.Web;
#else
using System.Collections.Specialized;
#endif
using DevExpress.ExpressApp;
using DevExpress.Xpo;
using DevExpress.Data.Filtering;
using Wxy.Xaf.ExcelExchange;

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// 默认Excel导入服务实现
    /// </summary>
    public class DefaultExcelImportService : IDialogExcelImportService
    {
        /// <summary>
        /// 从URI解析对象类型
        /// </summary>
        /// <param name="uri">URI</param>
        /// <returns>对象类型</returns>
        public Type ResolveObjectTypeFromUri(Uri uri)
        {
            try
            {
#if NET8_0_OR_GREATER
                var query = HttpUtility.ParseQueryString(uri.Query);
                var objectTypeName = query["objectType"];
#else
                // .NET Framework兼容实现
                var objectTypeName = ParseQueryParameter(uri.Query, "objectType");
#endif
                
                if (string.IsNullOrEmpty(objectTypeName))
                    return null;

                // 尝试从当前程序集中查找类型
                var assemblies = AppDomain.CurrentDomain.GetAssemblies();
                
                foreach (var assembly in assemblies)
                {
                    try
                    {
                        // 尝试完全限定名
                        var type = assembly.GetType(objectTypeName);
                        if (type != null)
                            return type;

                        // 尝试简单名称匹配
                        var types = assembly.GetTypes().Where(t => t.Name == objectTypeName).ToArray();
                        if (types.Length == 1)
                            return types[0];
                        
                        // 如果有多个匹配，优先选择XPObject的子类
                        var xpoTypes = types.Where(t => typeof(XPObject).IsAssignableFrom(t)).ToArray();
                        if (xpoTypes.Length == 1)
                            return xpoTypes[0];
                    }
                    catch
                    {
                        // 忽略程序集加载错误
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

#if !NET8_0_OR_GREATER
        /// <summary>
        /// 解析查询参数（.NET Framework兼容实现）
        /// </summary>
        /// <param name="query">查询字符串</param>
        /// <param name="parameterName">参数名</param>
        /// <returns>参数值</returns>
        private string ParseQueryParameter(string query, string parameterName)
        {
            if (string.IsNullOrEmpty(query))
                return null;

            if (query.StartsWith("?"))
                query = query.Substring(1);

            var pairs = query.Split('&');
            foreach (var pair in pairs)
            {
                var keyValue = pair.Split('=');
                if (keyValue.Length == 2 && keyValue[0] == parameterName)
                {
                    return Uri.UnescapeDataString(keyValue[1]);
                }
            }

            return null;
        }
#endif

        /// <summary>
        /// 解析导入模式
        /// </summary>
        /// <param name="modeString">模式字符串</param>
        /// <returns>导入模式</returns>
        public XpoExcelImportMode ParseImportMode(string modeString)
        {
            if (string.IsNullOrEmpty(modeString))
                return XpoExcelImportMode.CreateOrUpdate;

            switch (modeString.ToLower())
            {
                case "createonly":
                    return XpoExcelImportMode.CreateOnly;
                case "updateonly":
                    return XpoExcelImportMode.UpdateOnly;
                case "createorupdate":
                    return XpoExcelImportMode.CreateOrUpdate;
                case "replace":
                    return XpoExcelImportMode.Replace;
                default:
                    return XpoExcelImportMode.CreateOrUpdate;
            }
        }

        /// <summary>
        /// 从Excel流导入数据
        /// </summary>
        /// <param name="stream">Excel文件流</param>
        /// <param name="objectType">对象类型</param>
        /// <param name="options">导入选项</param>
        /// <param name="application">XAF应用程序实例</param>
        /// <returns>导入结果</returns>
        public XpoExcelImportResult ImportFromExcelStream(Stream stream, Type objectType, XpoExcelImportOptions options, XafApplication application)
        {
            var result = new XpoExcelImportResult();
            var debugInfo = new StringBuilder();

            try
            {
                // 将Excel流转换为CSV数据
                string csvData = ConvertExcelStreamToCsv(stream);

                if (string.IsNullOrEmpty(csvData))
                {
                    result.Errors.Add(new XpoExcelImportError
                    {
                        RowNumber = 0,
                        FieldName = "File",
                        ErrorMessage = "无法读取Excel文件或文件为空",
                        OriginalValue = ""
                    });
                    return result;
                }

                var importResult = ProcessCsvData(csvData, objectType, options, application, debugInfo);

                // 将调试信息添加到结果中
                if (importResult.Errors.Count > 0)
                {
                    importResult.Errors[0].ErrorMessage = $"[调试信息]\r\n{debugInfo.ToString()}\r\n[错误信息]\r\n{importResult.Errors[0].ErrorMessage}";
                }
                else if (importResult.SuccessCount > 0)
                {
                    importResult.Errors.Add(new XpoExcelImportError
                    {
                        RowNumber = 0,
                        FieldName = "Debug",
                        ErrorMessage = $"[调试信息]\r\n{debugInfo.ToString()}",
                        OriginalValue = ""
                    });
                }

                return importResult;
            }
            catch (Exception ex)
            {
                result.Errors.Add(new XpoExcelImportError
                {
                    RowNumber = 0,
                    FieldName = "File",
                    ErrorMessage = $"Excel导入失败: {ex.Message}\r\n[调试信息]\r\n{debugInfo.ToString()}",
                    OriginalValue = ""
                });
                return result;
            }
        }
        
        public XpoExcelImportResult ImportFromCsvStream(Stream stream, Type objectType, XpoExcelImportOptions options, XafApplication application)
        {
            var result = new XpoExcelImportResult();
            var debugInfo = new StringBuilder();

            try
            {
                // 检测编码：尝试多种编码方式读取CSV
                Encoding detectedEncoding = DetectCsvEncoding(stream, debugInfo);
                debugInfo.AppendLine($"检测到文件编码: {detectedEncoding.EncodingName}");

                // 重置流位置到开始
                stream.Position = 0;

                // 使用检测到的编码读取 CSV
                using (var reader = new StreamReader(stream, detectedEncoding))
                {
                    string csvData = reader.ReadToEnd();

                    debugInfo.AppendLine($"CSV文件读取完成，内容长度: {csvData?.Length ?? 0}");
                    if (!string.IsNullOrEmpty(csvData))
                    {
                        // 输出前200个字符用于调试
                        var preview = csvData.Length > 200 ? csvData.Substring(0, 200) : csvData;
                        debugInfo.AppendLine($"CSV内容预览: {preview}");
                    }

                    if (string.IsNullOrEmpty(csvData))
                    {
                        result.Errors.Add(new XpoExcelImportError
                        {
                            RowNumber = 0,
                            FieldName = "File",
                            ErrorMessage = "无法读取CSV文件或文件为空",
                            OriginalValue = ""
                        });
                        return result;
                    }

                    var importResult = ProcessCsvData(csvData, objectType, options, application, debugInfo);

                    // 将调试信息添加到第一个错误中（如果有的话），以便用户能看到
                    if (importResult.Errors.Count > 0)
                    {
                        importResult.Errors[0].ErrorMessage = $"[调试信息]\r\n{debugInfo.ToString()}\r\n[错误信息]\r\n{importResult.Errors[0].ErrorMessage}";
                    }
                    else if (importResult.SuccessCount > 0)
                    {
                        // 成功时也返回调试信息（通过添加一个信息性"错误"）
                        importResult.Errors.Add(new XpoExcelImportError
                        {
                            RowNumber = 0,
                            FieldName = "Debug",
                            ErrorMessage = $"[调试信息]\r\n{debugInfo.ToString()}",
                            OriginalValue = ""
                        });
                    }

                    return importResult;
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add(new XpoExcelImportError
                {
                    RowNumber = 0,
                    FieldName = "File",
                    ErrorMessage = $"CSV导入失败: {ex.Message}\r\n[调试信息]\r\n{debugInfo.ToString()}",
                    OriginalValue = ""
                });
                return result;
            }
        }
        
        /// <summary>
        /// 检测CSV文件的编码
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="debugInfo">调试信息</param>
        /// <returns>检测到的编码</returns>
        private Encoding DetectCsvEncoding(Stream stream, StringBuilder debugInfo)
        {
            // 记住当前位置
            long originalPosition = stream.Position;

            try
            {
                // 读取前1000个字节用于编码检测
                var buffer = new byte[1000];
                int bytesRead = stream.Read(buffer, 0, buffer.Length);
                stream.Position = originalPosition;

                if (bytesRead == 0)
                {
                    return Encoding.UTF8;
                }

                // 编码优先级列表：从最可能到最不可能
                var encodings = new List<Encoding>
                {
                    Encoding.UTF8,
                    Encoding.GetEncoding("GBK"),
                    Encoding.GetEncoding("GB2312"),
                    Encoding.GetEncoding("GB18030")
                };

                Encoding bestEncoding = Encoding.UTF8;
                int bestScore = int.MinValue;

                foreach (var encoding in encodings)
                {
                    int score = 0;
                    try
                    {
                        var testString = encoding.GetString(buffer, 0, bytesRead);

                        // 检查是否包含替换字符（表示编码不匹配）
                        int replacementCharCount = 0;
                        for (int i = 0; i < testString.Length && i < 100; i++)
                        {
                            if (testString[i] == '\ufffd')
                            {
                                replacementCharCount++;
                            }
                        }

                        // 检查中文字符范围（CJK统一汉字）
                        int chineseCharCount = 0;
                        for (int i = 0; i < testString.Length && i < 200; i++)
                        {
                            char c = testString[i];
                            if (c >= 0x4E00 && c <= 0x9FFF)
                            {
                                chineseCharCount++;
                            }
                        }

                        // 检查可打印ASCII字符比例
                        int printableAsciiCount = 0;
                        int totalChars = Math.Min(testString.Length, 200);
                        for (int i = 0; i < totalChars; i++)
                        {
                            char c = testString[i];
                            if ((c >= 32 && c <= 126) || c == '\r' || c == '\n' || c == '\t')
                            {
                                printableAsciiCount++;
                            }
                        }

                        // 计算得分：
                        // - 替换字符越少越好（-100分/个）
                        // - 中文字符越多越好（+10分/个）
                        // - 可打印ASCII比例越高越好（+1分/个）
                        score = (replacementCharCount * -100) + (chineseCharCount * 10) + printableAsciiCount;

                        debugInfo.AppendLine($"编码检测: {encoding.EncodingName} - 得分={score}, 替换字符={replacementCharCount}, 中文字符={chineseCharCount}, 可打印ASCII={printableAsciiCount}");

                        if (score > bestScore)
                        {
                            bestScore = score;
                            bestEncoding = encoding;
                        }
                    }
                    catch
                    {
                        // 编码不支持，跳过
                        debugInfo.AppendLine($"编码检测: {encoding.EncodingName} - 不支持，跳过");
                    }
                }

                debugInfo.AppendLine($"编码检测: 选择最佳编码 -> {bestEncoding.EncodingName}");
                return bestEncoding;
            }
            catch
            {
                // 检测失败，使用UTF-8
                return Encoding.UTF8;
            }
        }

        private XpoExcelImportResult ProcessCsvData(string csvData, Type objectType, XpoExcelImportOptions options, XafApplication application, StringBuilder debugInfo)
        {
            var result = new XpoExcelImportResult();

            try
            {
                // 解析CSV数据
                var csvRows = ParseCsvData(csvData);
                if (csvRows.Count == 0)
                {
                    result.Errors.Add(new XpoExcelImportError
                    {
                        RowNumber = 0,
                        FieldName = "Data",
                        ErrorMessage = "CSV数据为空",
                        OriginalValue = ""
                    });
                    return result;
                }

                // 获取表头和数据行
                var headers = csvRows[0];
                var dataRows = csvRows.Skip(1).ToList();

                debugInfo.AppendLine($"CSV解析完成: 表头=[{string.Join(", ", headers)}], 数据行数={dataRows.Count}");

                // 获取可导入的属性（带列名映射）
                var propertyMappings = GetImportableProperties(objectType);
                debugInfo.AppendLine($"可导入属性数: {propertyMappings.Count}");
                foreach (var mapping in propertyMappings)
                {
                    debugInfo.AppendLine($"  属性: {mapping.Key.Name} -> 列名: \"{mapping.Value}\"");
                }

                // 创建ObjectSpace
                using (var objectSpace = application.CreateObjectSpace())
                {
                    int rowIndex = 2; // 从第2行开始（第1行是表头）

                    foreach (var dataRow in dataRows)
                    {
                        debugInfo.AppendLine($"处理行 {rowIndex}: 数据=[{string.Join(", ", dataRow)}]");

                        try
                        {
                            // 根据导入模式处理记录
                            bool success = ProcessDataRow(objectSpace, objectType, headers, dataRow, propertyMappings, options, rowIndex, result, debugInfo);

                            if (success)
                            {
                                result.SuccessCount++;
                            }
                            else
                            {
                                result.FailedCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            debugInfo.AppendLine($"  行处理异常: {ex.Message}");
                            result.FailedCount++;
                            result.Errors.Add(new XpoExcelImportError
                            {
                                RowNumber = rowIndex,
                                FieldName = "Row",
                                ErrorMessage = ex.Message,
                                OriginalValue = string.Join(",", dataRow)
                            });
                        }

                        rowIndex++;
                    }

                    debugInfo.AppendLine($"导入完成: 成功={result.SuccessCount}, 失败={result.FailedCount}");

                    // 在保存前检查是否有待保存的更改
                    debugInfo.AppendLine($"准备保存更改到数据库...");
                    try
                    {
                        // 尝试获取 ObjectSpace 的详细信息
                        debugInfo.AppendLine($"ObjectSpace 类型: {objectSpace.GetType().FullName}");

                        debugInfo.AppendLine($"准备提交更改...");
                    }
                    catch
                    {
                        debugInfo.AppendLine($"检查 ObjectSpace 状态时出错");
                    }

                    // 保存更改
                    objectSpace.CommitChanges();
                    debugInfo.AppendLine($"更改已提交到数据库");
                }
            }
            catch (Exception ex)
            {
                debugInfo.AppendLine($"导入过程异常: {ex.Message}");
                result.Errors.Add(new XpoExcelImportError
                {
                    RowNumber = 0,
                    FieldName = "Import",
                    ErrorMessage = $"导入过程发生错误: {ex.Message}",
                    OriginalValue = ""
                });
            }

            return result;
        }

        /// <summary>
        /// 将Excel流转换为CSV数据
        /// </summary>
        /// <param name="stream">Excel流</param>
        /// <returns>CSV数据</returns>
        private string ConvertExcelStreamToCsv(Stream stream)
        {
            try
            {
                // 简单实现：假设输入的是CSV文件
                // 在实际项目中，这里应该使用Excel读取库如EPPlus或ClosedXML
                var debugInfo = new StringBuilder();
                Encoding detectedEncoding = DetectCsvEncoding(stream, debugInfo);

                // 重置流位置
                stream.Position = 0;

                using (var reader = new StreamReader(stream, detectedEncoding))
                {
                    return reader.ReadToEnd();
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// 解析CSV数据
        /// </summary>
        /// <param name="csvData">CSV数据</param>
        /// <returns>行数据列表</returns>
        private List<List<string>> ParseCsvData(string csvData)
        {
            var result = new List<List<string>>();
            
            if (string.IsNullOrEmpty(csvData))
                return result;

            var lines = csvData.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var line in lines)
            {
                var row = ParseCsvLine(line);
                if (row.Count > 0)
                {
                    result.Add(row);
                }
            }

            return result;
        }

        /// <summary>
        /// 解析CSV行
        /// </summary>
        /// <param name="line">CSV行</param>
        /// <returns>字段列表</returns>
        private List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            if (string.IsNullOrEmpty(line))
                return result;

            bool inQuotes = false;
            string currentField = string.Empty;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        currentField += '"';
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(currentField);
                    currentField = string.Empty;
                }
                else
                {
                    currentField += c;
                }
            }

            result.Add(currentField);
            return result;
        }

        /// <summary>
        /// 获取可导入的属性（带列名映射）
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <returns>属性和列名的映射列表</returns>
        private Dictionary<PropertyInfo, string> GetImportableProperties(Type objectType)
        {
            var result = new Dictionary<PropertyInfo, string>();
            var excludedFields = new[] { "This", "Loading", "ClassInfo", "Session", "IsLoading", "IsDeleted", "Oid", "GCRecord", "OptimisticLockField" };

            var properties = objectType
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanRead && p.CanWrite && !excludedFields.Contains(p.Name));

            foreach (var property in properties)
            {
                // 获取 ExcelFieldAttribute 特性
                var excelFieldAttr = (ExcelFieldAttribute)Attribute.GetCustomAttribute(property, typeof(ExcelFieldAttribute));

                // 如果有 ExcelFieldAttribute 特性且启用了导入，使用其 ColumnName
                if (excelFieldAttr != null && excelFieldAttr.EnabledImport)
                {
                    result[property] = excelFieldAttr.ColumnName ?? property.Name;
                }
                else if (excelFieldAttr == null)
                {
                    // 如果没有特性，使用属性名作为列名
                    result[property] = property.Name;
                }
                // 如果有特性但禁用了导入，则不添加到映射中
            }

            return result;
        }

        /// <summary>
        /// 处理数据行
        /// </summary>
        /// <param name="objectSpace">对象空间</param>
        /// <param name="objectType">对象类型</param>
        /// <param name="headers">表头</param>
        /// <param name="dataRow">数据行</param>
        /// <param name="propertyMappings">属性和列名的映射</param>
        /// <param name="options">导入选项</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="result">结果对象</param>
        /// <returns>是否成功</returns>
        private bool ProcessDataRow(IObjectSpace objectSpace, Type objectType, List<string> headers, List<string> dataRow,
            Dictionary<PropertyInfo, string> propertyMappings, XpoExcelImportOptions options, int rowIndex, XpoExcelImportResult result, StringBuilder debugInfo)
        {
            try
            {
                debugInfo.AppendLine($"  表头数={headers.Count}, 数据列数={dataRow.Count}");

                // 创建新对象
                var newObject = objectSpace.CreateObject(objectType);
                debugInfo.AppendLine($"  创建对象: 请求类型={objectType.Name}, 实际类型={newObject?.GetType().Name}");

                if (newObject == null)
                {
                    debugInfo.AppendLine($"  错误: 创建对象返回 null");
                    result.Errors.Add(new XpoExcelImportError
                    {
                        RowNumber = rowIndex,
                        FieldName = "Object",
                        ErrorMessage = "创建对象失败：返回 null",
                        OriginalValue = string.Join(",", dataRow)
                    });
                    return false;
                }

                // 获取实际对象类型的属性（可能不同于请求的类型）
                var actualObjectType = newObject.GetType();
                var actualProperties = actualObjectType.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Where(p => p.CanRead && p.CanWrite)
                    .ToDictionary(p => p.Name, p => p);

                debugInfo.AppendLine($"  实际对象类型: {actualObjectType.FullName}");
                debugInfo.AppendLine($"  实际对象可写属性数: {actualProperties.Count}");

                // 设置属性值
                for (int i = 0; i < headers.Count && i < dataRow.Count; i++)
                {
                    var header = headers[i];
                    var value = dataRow[i];

                    debugInfo.AppendLine($"  列 {i}: 表头=\"{header}\", 值=\"{value}\"");

                    // 根据列名查找属性
                    var propertyMapping = propertyMappings.FirstOrDefault(p =>
                        string.Equals(p.Value, header, StringComparison.OrdinalIgnoreCase));

                    if (propertyMapping.Key != null)
                    {
                        var requestedProperty = propertyMapping.Key;
                        debugInfo.AppendLine($"    匹配到属性: {requestedProperty.Name}");

                        // 尝试在实际对象类型中找到同名属性
                        if (!actualProperties.ContainsKey(requestedProperty.Name))
                        {
                            debugInfo.AppendLine($"    警告: 实际对象类型中找不到属性 {requestedProperty.Name}");
                            continue;
                        }

                        var actualProperty = actualProperties[requestedProperty.Name];
                        debugInfo.AppendLine($"    使用实际属性: {actualProperty.Name}, 类型={actualProperty.PropertyType.Name}");

                        try
                        {
                            var convertedValue = ConvertValue(value, actualProperty.PropertyType, requestedProperty, objectSpace);
                            actualProperty.SetValue(newObject, convertedValue);

                            // 验证值是否设置成功
                            var verifyValue = actualProperty.GetValue(newObject);
                            debugInfo.AppendLine($"    属性设置成功: {actualProperty.Name} = {convertedValue} (验证: {verifyValue})");
                        }
                        catch (Exception ex)
                        {
                            debugInfo.AppendLine($"    属性设置失败: {ex.Message}");
                            result.Errors.Add(new XpoExcelImportError
                            {
                                RowNumber = rowIndex,
                                FieldName = header,
                                ErrorMessage = $"值转换失败: {ex.Message}",
                                OriginalValue = value
                            });
                        }
                    }
                    else
                    {
                        debugInfo.AppendLine($"    未找到匹配的属性");
                    }
                }

                // 输出对象的最终状态用于调试
                debugInfo.AppendLine($"  对象创建完成，类型={newObject.GetType().Name}");
                try
                {
                    // 尝试获取关键属性的值进行验证
                    foreach (var mapping in propertyMappings)
                    {
                        try
                        {
                            if (actualProperties.ContainsKey(mapping.Key.Name))
                            {
                                var actualProp = actualProperties[mapping.Key.Name];
                                var propValue = actualProp.GetValue(newObject);
                                debugInfo.AppendLine($"    最终值: {actualProp.Name} = {propValue ?? "null"}");
                            }
                        }
                        catch
                        {
                            // 获取属性值失败,跳过
                        }
                    }
                }
                catch
                {
                    // 收集调试信息失败,跳过
                }

                return true;
            }
            catch (Exception ex)
            {
                debugInfo.AppendLine($"  行处理失败: {ex.Message}");
                result.Errors.Add(new XpoExcelImportError
                {
                    RowNumber = rowIndex,
                    FieldName = "Object",
                    ErrorMessage = $"创建对象失败: {ex.Message}",
                    OriginalValue = string.Join(",", dataRow)
                });
                return false;
            }
        }

        /// <summary>
        /// 转换值类型
        /// </summary>
        /// <param name="value">原始值</param>
        /// <param name="targetType">目标类型</param>
        /// <param name="property">属性信息（用于获取特性）</param>
        /// <param name="objectSpace">对象空间（用于查找引用对象）</param>
        /// <returns>转换后的值</returns>
        private object ConvertValue(string value, Type targetType, PropertyInfo property = null, IObjectSpace objectSpace = null)
        {
            if (string.IsNullOrEmpty(value))
            {
                if (targetType.IsGenericType && targetType.GetGenericTypeDefinition() == typeof(Nullable<>))
                {
                    return null;
                }
                if (targetType == typeof(string))
                {
                    return string.Empty;
                }
                if (targetType.IsValueType)
                {
                    return Activator.CreateInstance(targetType);
                }
                return null;
            }

            // 处理可空类型
            if (targetType.IsGenericType && targetType.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                targetType = Nullable.GetUnderlyingType(targetType);
            }

            // 检查是否为引用类型（XPO持久化对象）
            if (IsXpoPersistentObject(targetType) && property != null && objectSpace != null)
            {
                return ResolveReference(value, targetType, property, objectSpace);
            }

            // 特殊类型处理
            if (targetType == typeof(DateTime))
            {
                if (DateTime.TryParse(value, out DateTime dateValue))
                {
                    return dateValue;
                }
                return DateTime.MinValue;
            }

            if (targetType == typeof(bool))
            {
                value = value.ToLower();
                return value == "true" || value == "1" || value == "是" || value == "yes";
            }

            if (targetType.IsEnum)
            {
                // 对值进行 trim，去除前后空格
                var trimmedValue = value.Trim();

                // 获取枚举映射格式
                string enumFormat = null;
                if (property != null)
                {
                    var excelFieldAttr = (ExcelFieldAttribute)Attribute.GetCustomAttribute(property, typeof(ExcelFieldAttribute));
                    if (excelFieldAttr != null && !string.IsNullOrEmpty(excelFieldAttr.Format) && excelFieldAttr.Format.Contains("="))
                    {
                        enumFormat = excelFieldAttr.Format;
                    }
                }

                // 如果有映射格式，先尝试通过映射转换
                if (!string.IsNullOrEmpty(enumFormat))
                {
                    var mappings = ParseEnumMappings(enumFormat);
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

            // 通用转换
            return Convert.ChangeType(value, targetType);
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

        /// <summary>
        /// 检查类型是否为XPO持久化对象
        /// </summary>
        private bool IsXpoPersistentObject(Type type)
        {
            if (type == null || type == typeof(string))
                return false;

            // 检查是否继承自XPBaseObject
            var currentType = type;
            while (currentType != null && currentType != typeof(object))
            {
                if (currentType.Name == "XPBaseObject" || currentType.Name == "BaseObject" || currentType.Name == "XPLiteObject")
                    return true;

                currentType = currentType.BaseType;
            }

            // 或者检查是否实现了特定接口
            return type.GetInterface("IXPSimpleObject") != null;
        }

        /// <summary>
        /// 解析引用对象
        /// </summary>
        private object ResolveReference(string referenceValue, Type targetType, PropertyInfo property, IObjectSpace objectSpace)
        {
            try
            {
                // 获取ExcelFieldAttribute特性
                var excelFieldAttr = (ExcelFieldAttribute)Attribute.GetCustomAttribute(property, typeof(ExcelFieldAttribute));

                if (excelFieldAttr == null)
                {
                                        return null;
                }

                // 获取引用匹配字段（默认为"Oid"）
                string matchField = excelFieldAttr.ReferenceMatchField ?? "Oid";

                
                // 构建查询条件
                var criteria = CriteriaOperator.Parse($"[{matchField}] = ?", referenceValue);

                // 在ObjectSpace中查找对象
                var referencedObject = objectSpace.FindObject(targetType, criteria);

                if (referencedObject != null)
                {
                                        return referencedObject;
                }
                else
                {
                    
                    // 如果配置了ReferenceCreateIfNotExists，创建新对象
                    if (excelFieldAttr.ReferenceCreateIfNotExists)
                    {
                                                var newObject = objectSpace.CreateObject(targetType);

                        // 设置匹配字段的值
                        var matchProperty = targetType.GetProperty(matchField);
                        if (matchProperty != null)
                        {
                            // 根据匹配字段的类型进行简单转换（只处理基本类型，避免递归）
                            object matchFieldValue = null;
                            var propType = matchProperty.PropertyType;

                            // 处理可空类型
                            if (propType.IsGenericType && propType.GetGenericTypeDefinition() == typeof(Nullable<>))
                            {
                                propType = Nullable.GetUnderlyingType(propType);
                            }

                            if (propType == typeof(string))
                            {
                                matchFieldValue = referenceValue;
                            }
                            else if (propType == typeof(int) || propType == typeof(int?))
                            {
                                if (int.TryParse(referenceValue, out int intVal))
                                    matchFieldValue = intVal;
                            }
                            else if (propType == typeof(long) || propType == typeof(long?))
                            {
                                if (long.TryParse(referenceValue, out long longVal))
                                    matchFieldValue = longVal;
                            }
                            else if (propType == typeof(Guid) || propType == typeof(Guid?))
                            {
                                if (Guid.TryParse(referenceValue, out Guid guidVal))
                                    matchFieldValue = guidVal;
                            }

                            if (matchFieldValue != null)
                            {
                                matchProperty.SetValue(newObject, matchFieldValue);
                            }
                        }

                        return newObject;
                    }

                    return null;
                }
            }
            catch
            {
                                return null;
            }
        }
    }
}