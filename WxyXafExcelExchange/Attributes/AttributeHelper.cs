using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Wxy.Xaf.ExcelExchange;

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// ExcelField特性解析和验证辅助类
    /// </summary>
    public static class AttributeHelper
    {
        /// <summary>
        /// 获取类型的所有可导入属性
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <returns>可导入的属性列表</returns>
        public static List<PropertyInfo> GetImportableProperties(Type objectType)
        {
            var excludedFields = new[] { "This", "Loading", "ClassInfo", "Session", "IsLoading", "IsDeleted", "Oid", "GCRecord", "OptimisticLockField" };
            
            return objectType
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanRead && p.CanWrite && !excludedFields.Contains(p.Name))
                .Where(p => IsPropertyImportable(p))
                .OrderBy(p => p.Name)
                .ToList();
        }

        /// <summary>
        /// 获取类型的所有可导出属性
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <returns>可导出的属性列表</returns>
        public static List<PropertyInfo> GetExportableProperties(Type objectType)
        {
            var excludedFields = new[] { "This", "Loading", "ClassInfo", "Session", "IsLoading", "IsDeleted", "Oid", "GCRecord", "OptimisticLockField" };
            
            return objectType
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanRead && !excludedFields.Contains(p.Name))
                .Where(p => IsPropertyExportable(p))
                .OrderBy(p => p.Name)
                .ToList();
        }

        /// <summary>
        /// 检查属性是否可导入
        /// </summary>
        /// <param name="property">属性信息</param>
        /// <returns>是否可导入</returns>
        public static bool IsPropertyImportable(PropertyInfo property)
        {
            var excelField = property.GetCustomAttribute<ExcelFieldAttribute>();
            return excelField?.Import ?? true; // 默认可导入
        }

        /// <summary>
        /// 检查属性是否可导出
        /// </summary>
        /// <param name="property">属性信息</param>
        /// <returns>是否可导出</returns>
        public static bool IsPropertyExportable(PropertyInfo property)
        {
            var excelField = property.GetCustomAttribute<ExcelFieldAttribute>();
            return excelField?.Export ?? true; // 默认可导出
        }

        /// <summary>
        /// 获取属性的ExcelField特性
        /// </summary>
        /// <param name="property">属性信息</param>
        /// <returns>ExcelField特性，如果没有则返回null</returns>
        public static ExcelFieldAttribute GetExcelFieldAttribute(PropertyInfo property)
        {
            return property.GetCustomAttribute<ExcelFieldAttribute>();
        }

        /// <summary>
        /// 获取属性的有效列名
        /// </summary>
        /// <param name="property">属性信息</param>
        /// <returns>有效的列名</returns>
        public static string GetEffectiveColumnName(PropertyInfo property)
        {
            var excelField = GetExcelFieldAttribute(property);
            return excelField?.GetEffectiveColumnName(property.Name) ?? property.Name;
        }

        /// <summary>
        /// 检查属性是否为必填字段
        /// </summary>
        /// <param name="property">属性信息</param>
        /// <returns>是否必填</returns>
        public static bool IsRequired(PropertyInfo property)
        {
            var excelField = GetExcelFieldAttribute(property);
            return excelField?.Required ?? false;
        }

        /// <summary>
        /// 获取属性的默认值
        /// </summary>
        /// <param name="property">属性信息</param>
        /// <returns>默认值，如果没有则返回null</returns>
        public static string GetDefaultValue(PropertyInfo property)
        {
            var excelField = GetExcelFieldAttribute(property);
            return excelField?.DefaultValue;
        }

        /// <summary>
        /// 获取属性的自定义格式
        /// </summary>
        /// <param name="property">属性信息</param>
        /// <returns>自定义格式，如果没有则返回null</returns>
        public static string GetCustomFormat(PropertyInfo property)
        {
            var excelField = GetExcelFieldAttribute(property);
            return excelField?.Format;
        }

        /// <summary>
        /// 获取属性的关联对象查找键
        /// </summary>
        /// <param name="property">属性信息</param>
        /// <returns>查找键，如果没有则返回null</returns>
        public static string GetLookupKey(PropertyInfo property)
        {
            var excelField = GetExcelFieldAttribute(property);
            return excelField?.Key;
        }

        /// <summary>
        /// 检查属性是否有关联对象查找配置
        /// </summary>
        /// <param name="property">属性信息</param>
        /// <returns>是否有查找配置</returns>
        public static bool HasLookupKey(PropertyInfo property)
        {
            var excelField = GetExcelFieldAttribute(property);
            return excelField?.HasKeyLookup() ?? false;
        }

        /// <summary>
        /// 验证属性值是否符合要求
        /// </summary>
        /// <param name="property">属性信息</param>
        /// <param name="value">属性值</param>
        /// <param name="errorMessage">错误消息</param>
        /// <returns>是否验证通过</returns>
        public static bool ValidatePropertyValue(PropertyInfo property, string value, out string errorMessage)
        {
            errorMessage = null;
            
            var excelField = GetExcelFieldAttribute(property);
            if (excelField == null)
                return true;

            // 检查必填字段
            if (excelField.Required && string.IsNullOrWhiteSpace(value))
            {
                errorMessage = $"字段 '{GetEffectiveColumnName(property)}' 是必填字段，不能为空";
                return false;
            }

            return true;
        }

        /// <summary>
        /// 获取属性的最终值（考虑默认值）
        /// </summary>
        /// <param name="property">属性信息</param>
        /// <param name="originalValue">原始值</param>
        /// <returns>最终值</returns>
        public static string GetFinalValue(PropertyInfo property, string originalValue)
        {
            if (!string.IsNullOrWhiteSpace(originalValue))
                return originalValue;

            var defaultValue = GetDefaultValue(property);
            return !string.IsNullOrEmpty(defaultValue) ? defaultValue : originalValue;
        }

        /// <summary>
        /// 创建属性到列名的映射字典
        /// </summary>
        /// <param name="properties">属性列表</param>
        /// <returns>列名到属性的映射字典</returns>
        public static Dictionary<string, PropertyInfo> CreateColumnToPropertyMap(IEnumerable<PropertyInfo> properties)
        {
            var map = new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);
            
            foreach (var property in properties)
            {
                var columnName = GetEffectiveColumnName(property);
                map[columnName] = property;
            }
            
            return map;
        }
    }
}