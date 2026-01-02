using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.DC;
using Wxy.Xaf.ExcelExchange;
using Wxy.Xaf.ExcelExchange.Configuration;
using Wxy.Xaf.ExcelExchange.Enums;

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// Excel配置管理服务，负责解析和管理类和属性上的特性配置
    /// </summary>
    public class ConfigurationManager
    {
        private readonly ConcurrentDictionary<Type, ExcelConfiguration> _configurationCache = new ConcurrentDictionary<Type, ExcelConfiguration>();

        /// <summary>
        /// 获取对象类型的Excel配置
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <param name="objectTypeInfo">XAF对象类型信息（可选，用于获取本地化信息）</param>
        /// <returns>Excel配置</returns>
        public ExcelConfiguration GetConfiguration(Type objectType, ITypeInfo objectTypeInfo = null)
        {
            if (objectType == null)
                throw new ArgumentNullException(nameof(objectType));

            // 检查缓存
            if (_configurationCache.TryGetValue(objectType, out var cachedConfig))
            {
                return cachedConfig;
            }

            // 创建新配置
            var configuration = CreateConfiguration(objectType, objectTypeInfo);

            // 验证配置
            var validation = configuration.ValidateConfiguration();
            if (!validation.IsValid)
            {
                var errors = string.Join("; ", validation.Errors);
                throw new InvalidOperationException($"Excel配置验证失败 ({objectType.Name}): {errors}");
            }

            // 线程安全地缓存配置
            _configurationCache.TryAdd(objectType, configuration);

            return configuration;
        }

        /// <summary>
        /// 检查对象类型是否支持Excel导入导出
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <returns>是否支持</returns>
        public bool IsExcelSupported(Type objectType)
        {
            if (objectType == null) return false;

            try
            {
                // 检查是否有ExcelImportExportAttribute
                var classAttribute = objectType.GetCustomAttribute<ExcelImportExportAttribute>();
                if (classAttribute != null)
                {
                    return classAttribute.EnabledImport || classAttribute.EnabledExport;
                }

                // 检查是否有任何属性标记了ExcelFieldAttribute
                var properties = GetExcelProperties(objectType);
                return properties.Any(p => p.GetCustomAttribute<ExcelFieldAttribute>() != null);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 清除配置缓存
        /// </summary>
        /// <param name="objectType">要清除的对象类型，null表示清除所有</param>
        public void ClearCache(Type objectType = null)
        {
            if (objectType == null)
            {
                _configurationCache.Clear();
            }
            else
            {
                _configurationCache.TryRemove(objectType, out _);
            }
        }

        /// <summary>
        /// 获取所有已缓存的配置
        /// </summary>
        /// <returns>配置字典</returns>
        public Dictionary<Type, ExcelConfiguration> GetAllConfigurations()
        {
            return new Dictionary<Type, ExcelConfiguration>(_configurationCache);
        }

        #region 私有方法

        /// <summary>
        /// 创建Excel配置
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <param name="objectTypeInfo">XAF对象类型信息</param>
        /// <returns>Excel配置</returns>
        private ExcelConfiguration CreateConfiguration(Type objectType, ITypeInfo objectTypeInfo)
        {
            var configuration = new ExcelConfiguration
            {
                ObjectType = objectType
            };

            // 解析类级别配置
            configuration.ClassConfiguration = ParseClassConfiguration(objectType, objectTypeInfo);

            // 解析字段配置
            configuration.FieldConfigurations = ParseFieldConfigurations(objectType, objectTypeInfo);

            return configuration;
        }

        /// <summary>
        /// 解析类级别配置
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <param name="objectTypeInfo">XAF对象类型信息</param>
        /// <returns>类级别配置</returns>
        private ExcelImportExportAttribute ParseClassConfiguration(Type objectType, ITypeInfo objectTypeInfo)
        {
            // 优先从XAF TypeInfo获取
            var attribute = objectTypeInfo?.FindAttribute<ExcelImportExportAttribute>();
            
            // 如果没有，从类型直接获取
            if (attribute == null)
            {
                attribute = objectType.GetCustomAttribute<ExcelImportExportAttribute>();
            }

            // 如果仍然没有，创建默认配置
            if (attribute == null)
            {
                attribute = new ExcelImportExportAttribute();
            }

            // 设置默认工作表名称（如果未指定）
            if (string.IsNullOrEmpty(attribute.SheetName))
            {
                // 尝试从XAF获取本地化名称
                if (objectTypeInfo != null)
                {
                    try
                    {
                        // 这里可以集成XAF的本地化机制
                        // attribute.SheetName = objectTypeInfo.DisplayName;
                    }
                    catch
                    {
                        // 忽略错误，使用默认值
                    }
                }
                
                // 如果仍然为空，使用类名
                if (string.IsNullOrEmpty(attribute.SheetName))
                {
                    attribute.SheetName = objectType.Name;
                }
            }

            return attribute;
        }

        /// <summary>
        /// 解析字段配置
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <param name="objectTypeInfo">XAF对象类型信息</param>
        /// <returns>字段配置列表</returns>
        private List<ExcelFieldConfiguration> ParseFieldConfigurations(Type objectType, ITypeInfo objectTypeInfo)
        {
            var fieldConfigurations = new List<ExcelFieldConfiguration>();
            var properties = GetExcelProperties(objectType);

            // 首先记录每个属性的声明顺序（用于未明确指定 Order 的字段）
            var propertyDeclarationOrder = new Dictionary<string, int>();
            for (int i = 0; i < properties.Length; i++)
            {
                propertyDeclarationOrder[properties[i].Name] = i;
            }

            // 第一遍：解析所有字段特性，不设置默认值
            foreach (var property in properties)
            {
                var fieldAttribute = ParseFieldAttribute(property, objectTypeInfo);

                var fieldConfiguration = new ExcelFieldConfiguration
                {
                    PropertyInfo = property,
                    FieldAttribute = fieldAttribute
                };

                fieldConfigurations.Add(fieldConfiguration);
            }

            // 对字段配置进行排序
            // 规则：
            // 1. 有明确 Order 值的字段按 Order 值升序排列（Order值最小的在最左边）
            // 2. Order 值相同时，按代码声明顺序排列
            // 3. 没有明确 Order 值的字段按代码声明顺序排列，排在所有有 Order 值的字段之后
            var sortedFieldConfigurations = fieldConfigurations
                .OrderBy(f =>
                {
                    // 判断是否有明确的 Order 设置（SortOrder > 0 或 ColumnIndex > 0）
                    bool hasExplicitOrder = f.FieldAttribute.SortOrder > 0 || f.FieldAttribute.ColumnIndex > 0;

                    if (hasExplicitOrder)
                    {
                        // 使用明确设置的 SortOrder 或 ColumnIndex（优先使用 SortOrder）
                        int order = f.FieldAttribute.SortOrder > 0 ? f.FieldAttribute.SortOrder : f.FieldAttribute.ColumnIndex;
                        return order;
                    }
                    else
                    {
                        // 没有明确 Order 的字段，使用 int.MaxValue 确保排在所有有 Order 的字段之后
                        return int.MaxValue;
                    }
                })
                .ThenBy(f =>
                {
                    // 所有字段在相同 Order 值的情况下，都按代码声明顺序排列
                    // 这确保了：
                    // - 有明确 Order 值的字段，Order 相同时按声明顺序
                    // - 没有明确 Order 值的字段，按声明顺序排列
                    return propertyDeclarationOrder[f.PropertyInfo.Name];
                })
                .ToList();

            return sortedFieldConfigurations;
        }

        /// <summary>
        /// 解析字段特性
        /// </summary>
        /// <param name="property">属性信息</param>
        /// <param name="objectTypeInfo">XAF对象类型信息</param>
        /// <returns>字段特性</returns>
        private ExcelFieldAttribute ParseFieldAttribute(PropertyInfo property, ITypeInfo objectTypeInfo)
        {
            // 获取属性上的ExcelFieldAttribute
            var attribute = property.GetCustomAttribute<ExcelFieldAttribute>();

            // 如果没有特性，创建默认配置
            if (attribute == null)
            {
                attribute = new ExcelFieldAttribute();
            }

            // 设置默认列名（如果未指定）
            if (string.IsNullOrEmpty(attribute.ColumnName))
            {
                // 尝试从XAF获取本地化名称
                if (objectTypeInfo != null)
                {
                    try
                    {
                        var memberInfo = objectTypeInfo.FindMember(property.Name);
                        if (memberInfo != null)
                        {
                            // 这里可以集成XAF的本地化机制
                            // attribute.ColumnName = memberInfo.DisplayName;
                        }
                    }
                    catch
                    {
                        // 忽略错误，使用默认值
                    }
                }
                
                // 如果仍然为空，使用属性名
                if (string.IsNullOrEmpty(attribute.ColumnName))
                {
                    attribute.ColumnName = property.Name;
                }
            }

            // 设置默认验证错误消息
            if (string.IsNullOrEmpty(attribute.ValidationErrorMessage))
            {
                attribute.ValidationErrorMessage = attribute.GetEffectiveValidationErrorMessage(property.Name);
            }

            // 根据属性类型设置默认配置
            SetDefaultConfigurationByType(attribute, property);

            return attribute;
        }

        /// <summary>
        /// 根据属性类型设置默认配置
        /// </summary>
        /// <param name="attribute">字段特性</param>
        /// <param name="property">属性信息</param>
        private void SetDefaultConfigurationByType(ExcelFieldAttribute attribute, PropertyInfo property)
        {
            var propertyType = property.PropertyType;

            // 处理可空类型
            if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                propertyType = Nullable.GetUnderlyingType(propertyType);
            }

            // XPCollection 集合类型 - 支持导出但不支持导入
            // 检查属性类型名称（优先检查 propertyType 本身）
            bool isXPCollection = propertyType.Name.Contains("XPCollection") ||
                (propertyType.FullName != null && propertyType.FullName.Contains("XPCollection"));

            // 再次检查泛型类型定义的名称
            if (!isXPCollection && propertyType.IsGenericType)
            {
                var genericTypeDefinition = propertyType.GetGenericTypeDefinition();
                isXPCollection = genericTypeDefinition.Name.Contains("XPCollection") ||
                    genericTypeDefinition.FullName != null && genericTypeDefinition.FullName.Contains("XPCollection");
            }

            if (isXPCollection)
            {
                // XPCollection: 只支持导出，不支持导入
                attribute.EnabledImport = false;
                // 如果没有显式禁用导出，则默认启用
                if (attribute.EnabledExport) // 默认为 true，除非用户显式设置为 false
                {
                    // 保持 EnabledExport = true，支持导出
                }
                // 如果用户没有显式设置 CollectionExportFormat，使用默认值
                if (string.IsNullOrEmpty(attribute.CollectionExportFormat) || attribute.CollectionExportFormat == "Summary")
                {
                    attribute.CollectionExportFormat = "Summary";
                }
                // XPCollection 不隐藏，允许显示
                // 不立即返回，继续下面的处理
            }

            // 根据类型设置默认格式和对齐方式
            if (propertyType == typeof(DateTime))
            {
                if (string.IsNullOrEmpty(attribute.Format))
                {
                    attribute.Format = "yyyy-MM-dd HH:mm:ss";
                }
            }
            else if (propertyType == typeof(decimal) || propertyType == typeof(double) || propertyType == typeof(float))
            {
                attribute.Alignment = ExcelAlignment.Right;
                if (string.IsNullOrEmpty(attribute.Format))
                {
                    attribute.Format = "0.00";
                }
            }
            else if (propertyType == typeof(int) || propertyType == typeof(long) || propertyType == typeof(short))
            {
                attribute.Alignment = ExcelAlignment.Right;
            }
            else if (propertyType == typeof(bool))
            {
                attribute.Alignment = ExcelAlignment.Center;
            }

            // 检查是否为主键或系统字段
            if (IsSystemField(property))
            {
                attribute.EnabledImport = false;
                attribute.Hidden = true;
            }

            // 检查是否为只读属性
            if (!property.CanWrite && !isXPCollection) // XPCollection 是只读的，但允许导出
            {
                attribute.EnabledImport = false;
            }
        }

        /// <summary>
        /// 获取Excel相关的属性
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <returns>属性列表</returns>
        private PropertyInfo[] GetExcelProperties(Type objectType)
        {
            // 排除XPO内置字段和系统字段
            var excludedFields = new[]
            {
                "This", "Loading", "ClassInfo", "Session", "IsLoading", "IsDeleted",
                "Oid", "GCRecord", "OptimisticLockField", "ObjectType"
            };

            // 保持属性声明顺序,不按名称排序
            return objectType
                .GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
                .Where(p => p.CanRead && !excludedFields.Contains(p.Name))
                .ToArray();
        }

        /// <summary>
        /// 检查是否为系统字段
        /// </summary>
        /// <param name="property">属性信息</param>
        /// <returns>是否为系统字段</returns>
        private bool IsSystemField(PropertyInfo property)
        {
            var systemFields = new[] 
            { 
                "Oid", "GCRecord", "OptimisticLockField", "ObjectType",
                "This", "Loading", "ClassInfo", "Session", "IsLoading", "IsDeleted"
            };

            return systemFields.Contains(property.Name) || 
                   property.Name.StartsWith("_") ||
                   property.GetCustomAttributes().Any(attr => attr.GetType().Name.Contains("Key"));
        }

        #endregion
    }
}