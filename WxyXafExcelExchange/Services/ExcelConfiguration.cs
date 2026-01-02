using System;
using System.Collections.Generic;
using System.Linq;
using Wxy.Xaf.ExcelExchange;
using Wxy.Xaf.ExcelExchange.Enums;

namespace Wxy.Xaf.ExcelExchange.Configuration
{
    /// <summary>
    /// Excel配置信息类，包含类级别和属性级别的完整配置
    /// </summary>
    public class ExcelConfiguration
    {
        #region 属性

        /// <summary>
        /// 对象类型
        /// </summary>
        public Type ObjectType { get; set; }

        /// <summary>
        /// 类级别配置
        /// </summary>
        public ExcelImportExportAttribute ClassConfiguration { get; set; }

        /// <summary>
        /// 字段配置列表
        /// </summary>
        public List<ExcelFieldConfiguration> FieldConfigurations { get; set; } = new List<ExcelFieldConfiguration>();

        #endregion

        #region 便捷属性

        /// <summary>
        /// 是否启用导入
        /// </summary>
        public bool IsImportEnabled => ClassConfiguration?.EnabledImport ?? false;

        /// <summary>
        /// 是否启用导出
        /// </summary>
        public bool IsExportEnabled => ClassConfiguration?.EnabledExport ?? false;

        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName => ClassConfiguration?.GetEffectiveSheetName(ObjectType) ?? ObjectType?.Name;

        /// <summary>
        /// 导入字段配置
        /// </summary>
        public List<ExcelFieldConfiguration> ImportFields => 
            FieldConfigurations?.Where(f => f.FieldAttribute.EnabledImport).OrderBy(f => f.EffectiveSortOrder).ToList() ?? new List<ExcelFieldConfiguration>();

        /// <summary>
        /// 导出字段配置
        /// 字段已由 ConfigurationManager 按以下规则排序:
        /// 1. 有明确 SortOrder/ColumnIndex 的字段按设置值排序
        /// 2. 没有明确 Order 的字段按属性声明顺序排列
        /// </summary>
        public List<ExcelFieldConfiguration> ExportFields =>
            FieldConfigurations?.Where(f => f.FieldAttribute.EnabledExport && !f.FieldAttribute.Hidden).ToList() ?? new List<ExcelFieldConfiguration>();

        /// <summary>
        /// 所有字段配置（包括隐藏字段）
        /// </summary>
        public List<ExcelFieldConfiguration> AllExportFields => 
            FieldConfigurations?.Where(f => f.FieldAttribute.EnabledExport).OrderBy(f => f.EffectiveSortOrder).ToList() ?? new List<ExcelFieldConfiguration>();

        #endregion

        #region 方法

        /// <summary>
        /// 根据属性名获取字段配置
        /// </summary>
        /// <param name="propertyName">属性名</param>
        /// <returns>字段配置</returns>
        public ExcelFieldConfiguration GetFieldConfiguration(string propertyName)
        {
            return FieldConfigurations?.FirstOrDefault(f => f.PropertyInfo.Name == propertyName);
        }

        /// <summary>
        /// 根据列名获取字段配置
        /// </summary>
        /// <param name="columnName">列名</param>
        /// <returns>字段配置</returns>
        public ExcelFieldConfiguration GetFieldConfigurationByColumnName(string columnName)
        {
            return FieldConfigurations?.FirstOrDefault(f => f.EffectiveColumnName.Equals(columnName, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// 验证配置的完整性
        /// </summary>
        /// <returns>验证结果</returns>
        public (bool IsValid, List<string> Errors) ValidateConfiguration()
        {
            var errors = new List<string>();

            // 验证类级别配置
            if (ClassConfiguration != null)
            {
                var classValidation = ClassConfiguration.ValidateConfiguration();
                if (!classValidation.IsValid)
                {
                    errors.Add($"类配置错误: {classValidation.ErrorMessage}");
                }
            }

            // 验证字段配置
            foreach (var fieldConfig in FieldConfigurations)
            {
                var fieldValidation = fieldConfig.FieldAttribute.ValidateConfiguration();
                if (!fieldValidation.IsValid)
                {
                    errors.Add($"字段 {fieldConfig.PropertyInfo.Name} 配置错误: {fieldValidation.ErrorMessage}");
                }
            }

            // 检查列名重复
            var columnNames = FieldConfigurations
                .Where(f => f.FieldAttribute.EnabledExport || f.FieldAttribute.EnabledImport)
                .GroupBy(f => f.EffectiveColumnName)
                .Where(g => g.Count() > 1)
                .Select(g => g.Key);

            foreach (var duplicateColumnName in columnNames)
            {
                errors.Add($"列名重复: {duplicateColumnName}");
            }

            // 检查列索引重复
            var columnIndexes = FieldConfigurations
                .Where(f => (f.FieldAttribute.EnabledExport || f.FieldAttribute.EnabledImport) && f.FieldAttribute.ColumnIndex > 0)
                .GroupBy(f => f.FieldAttribute.ColumnIndex)
                .Where(g => g.Count() > 1)
                .Select(g => g.Key);

            foreach (var duplicateIndex in columnIndexes)
            {
                errors.Add($"列索引重复: {duplicateIndex}");
            }

            return (errors.Count == 0, errors);
        }

        /// <summary>
        /// 获取配置摘要信息
        /// </summary>
        /// <returns>配置摘要</returns>
        public string GetConfigurationSummary()
        {
            var summary = new List<string>
            {
                $"对象类型: {ObjectType?.Name}",
                $"工作表名称: {SheetName}",
                $"导入启用: {IsImportEnabled}",
                $"导出启用: {IsExportEnabled}",
                $"导入字段数: {ImportFields.Count}",
                $"导出字段数: {ExportFields.Count}"
            };

            if (ClassConfiguration != null)
            {
                summary.Add($"重复策略: {ClassConfiguration.ImportDuplicateStrategy}");
                summary.Add($"验证模式: {ClassConfiguration.ValidationMode}");
                summary.Add($"自动保存: {ClassConfiguration.AutoSaveAfterImport}");
            }

            return string.Join("\n", summary);
        }

        #endregion
    }

    /// <summary>
    /// Excel字段配置信息
    /// </summary>
    public class ExcelFieldConfiguration
    {
        /// <summary>
        /// 属性信息
        /// </summary>
        public System.Reflection.PropertyInfo PropertyInfo { get; set; }

        /// <summary>
        /// 字段特性
        /// </summary>
        public ExcelFieldAttribute FieldAttribute { get; set; }

        /// <summary>
        /// 有效的列名
        /// </summary>
        public string EffectiveColumnName => FieldAttribute?.GetEffectiveColumnName(PropertyInfo?.Name) ?? PropertyInfo?.Name;

        /// <summary>
        /// 有效的排序顺序
        /// </summary>
        public int EffectiveSortOrder
        {
            get
            {
                if (FieldAttribute?.ColumnIndex > 0)
                {
                    return FieldAttribute.ColumnIndex * 1000; // 列索引优先级更高
                }
                return FieldAttribute?.SortOrder ?? 0;
            }
        }

        /// <summary>
        /// 是否为关联对象属性
        /// </summary>
        public bool IsReferenceProperty
        {
            get
            {
                if (PropertyInfo?.PropertyType == null) return false;
                
                // 检查是否为XPO对象类型
                var type = PropertyInfo.PropertyType;
                return !type.IsPrimitive && 
                       !type.IsEnum && 
                       type != typeof(string) && 
                       type != typeof(DateTime) && 
                       type != typeof(decimal) && 
                       type != typeof(Guid) &&
                       !type.IsGenericType;
            }
        }

        /// <summary>
        /// 是否有自定义转换器
        /// </summary>
        public bool HasCustomConverter => FieldAttribute?.HasCustomConverter() ?? false;

        /// <summary>
        /// 是否有验证规则
        /// </summary>
        public bool HasValidationRules => FieldAttribute?.HasValidationRules() ?? false;

        /// <summary>
        /// 获取字段配置摘要
        /// </summary>
        /// <returns>配置摘要</returns>
        public string GetConfigurationSummary()
        {
            var summary = new List<string>
            {
                $"属性: {PropertyInfo?.Name}",
                $"列名: {EffectiveColumnName}",
                $"类型: {PropertyInfo?.PropertyType?.Name}",
                $"导入: {FieldAttribute?.EnabledImport}",
                $"导出: {FieldAttribute?.EnabledExport}",
                $"必填: {FieldAttribute?.IsRequired}",
                $"排序: {EffectiveSortOrder}"
            };

            if (IsReferenceProperty)
            {
                summary.Add($"关联对象: {FieldAttribute?.ReferenceMatchField}");
                summary.Add($"自动创建: {FieldAttribute?.ReferenceCreateIfNotExists}");
            }

            if (HasCustomConverter)
            {
                summary.Add($"导入转换: {FieldAttribute?.ImportConvertMethod}");
                summary.Add($"导出转换: {FieldAttribute?.ExportConvertMethod}");
            }

            return string.Join(", ", summary);
        }
    }
}