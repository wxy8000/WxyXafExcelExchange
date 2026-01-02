using System;
using System.Collections.Generic;
using Wxy.Xaf.ExcelExchange.Enums;

namespace Wxy.Xaf.ExcelExchange
{
    /// <summary>
    /// 属性级别特性：控制单个XPO属性的Excel映射规则
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelFieldAttribute : Attribute
    {
        #region 基础开关与映射（必选核心）

        /// <summary>
        /// 是否启用该属性的导入，默认为true
        /// </summary>
        public bool EnabledImport { get; set; } = true;

        /// <summary>
        /// 是否启用该属性的导出，默认为true
        /// </summary>
        public bool EnabledExport { get; set; } = true;

        /// <summary>
        /// Excel列名（如属性名CustomerName → 列名"客户名称"）
        /// 默认使用属性的XAF DisplayName（本地化），如果没有则使用属性名
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// 列索引（从1开始），0表示按SortOrder自动排序
        /// 支持强制指定列位置（如"序号"固定在第1列）
        /// </summary>
        public int ColumnIndex { get; set; } = 0;

        /// <summary>
        /// 列显示顺序（数值越小越靠前），替代ColumnIndex的柔性排序
        /// </summary>
        public int SortOrder { get; set; } = 0;

        #endregion

        #region 导入验证（核心）

        /// <summary>
        /// 导入时该列是否必填，默认为false
        /// </summary>
        public bool IsRequired { get; set; } = false;

        /// <summary>
        /// 数据格式（日期/数值/枚举）
        /// 例："yyyy-MM-dd"、"0.00"、"Male=男;Female=女"
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 导入验证的正则表达式（如手机号^\d{11}$）
        /// </summary>
        public string ValidationRegex { get; set; }

        /// <summary>
        /// 验证失败时的提示文本，默认为"【{ColumnName}】格式错误"
        /// </summary>
        public string ValidationErrorMessage { get; set; }

        /// <summary>
        /// 导入时空值替换的默认值（如空字符串→0）
        /// </summary>
        public object ImportNullValueReplacement { get; set; }

        #endregion

        #region 导出格式化（核心）

        /// <summary>
        /// 导出时空值显示的文本（如null→"无"），默认为空字符串
        /// </summary>
        public string ExportNullValueDisplay { get; set; } = "";

        /// <summary>
        /// 列对齐方式，默认为Left
        /// </summary>
        public ExcelAlignment Alignment { get; set; } = ExcelAlignment.Left;

        /// <summary>
        /// 导出列宽（字符数），0表示自动调整
        /// </summary>
        public int ColumnWidth { get; set; } = 0;

        /// <summary>
        /// 导出时该列内容/表头是否加粗，默认为false
        /// </summary>
        public bool IsBold { get; set; } = false;

        /// <summary>
        /// 导出列是否隐藏（存在但不可见），默认为false
        /// </summary>
        public bool Hidden { get; set; } = false;

        /// <summary>
        /// Excel列批注（鼠标悬停显示）
        /// </summary>
        public string Comments { get; set; }

        #endregion

        #region 数据转换（进阶）

        /// <summary>
        /// 导入转换方法名（静态方法）
        /// 例："ConvertGenderFromExcel"
        /// </summary>
        public string ImportConvertMethod { get; set; }

        /// <summary>
        /// 导出转换方法名（静态方法）
        /// 例："ConvertGenderToExcel"
        /// </summary>
        public string ExportConvertMethod { get; set; }

        #endregion

        #region 关联对象处理（XPO专属核心）

        /// <summary>
        /// 关联对象的匹配字段，默认为"Oid"
        /// 如通过Company.Code匹配，而非Oid
        /// </summary>
        public string ReferenceMatchField { get; set; } = "Oid";

        /// <summary>
        /// 导入时关联对象匹配不到是否自动创建，默认为false
        /// </summary>
        public bool ReferenceCreateIfNotExists { get; set; } = false;

        /// <summary>
        /// 集合属性导出格式：
        /// - Summary：汇总为单个文本，如 "产品A x2; 产品B x3"
        /// - MultiSheet（推荐）：展开为独立的 Sheet（支持导入导出）
        /// - MultiRow：展开为多行（每行一个明细项）
        /// - Count：仅显示数量，如 "3条明细"
        /// 默认为 MultiSheet 以支持完整的导入导出功能
        /// </summary>
        public string CollectionExportFormat { get; set; } = "MultiSheet";

        /// <summary>
        /// 集合导出时的分隔符（Summary 模式），默认为 "; "
        /// </summary>
        public string CollectionDelimiter { get; set; } = "; ";

        /// <summary>
        /// 集合导出时显示的子属性名（逗号分隔），默认为第一个字符串属性
        /// 例如 "ProductName,Quantity" 显示为 "笔记本电脑 x2"
        /// </summary>
        public string CollectionDisplayProperties { get; set; } = "";

        /// <summary>
        /// 多 Sheet 模式下的子表 Sheet 名称
        /// 默认为 "{属性名}明细"，例如 "订单明细"
        /// </summary>
        public string DetailSheetName { get; set; } = "";

        /// <summary>
        /// 关联字段在子表中的列名（用于导入时关联主表）
        /// 默认为 "关联主表记录"，可以自定义如 "订单编号"
        /// </summary>
        public string RelationFieldName { get; set; } = "关联主表记录";

        /// <summary>
        /// 是否在导入时清空现有子表数据后再添加新数据
        /// 默认为 true（替换模式），设为 false 则追加数据
        /// </summary>
        public bool ClearExistingDetailsOnImport { get; set; } = true;

        #endregion

        #region 扩展参数（可选）

        /// <summary>
        /// 导出时属性为空则不显示该列，默认为false
        /// </summary>
        public bool IgnoreEmptyValue { get; set; } = false;

        /// <summary>
        /// 导出单元格背景色（如"#F5F5F5"）
        /// </summary>
        public string CellColor { get; set; }

        /// <summary>
        /// 导出模板中该列是否锁定（保护工作表），默认为false
        /// </summary>
        public bool IsLocked { get; set; } = false;

        #endregion

        #region 向后兼容属性

        /// <summary>
        /// 是否参与导入操作（向后兼容）
        /// </summary>
        public bool Import
        {
            get => EnabledImport;
            set => EnabledImport = value;
        }

        /// <summary>
        /// 是否参与导出操作（向后兼容）
        /// </summary>
        public bool Export
        {
            get => EnabledExport;
            set => EnabledExport = value;
        }

        /// <summary>
        /// 用于查找关联对象的键属性名称（向后兼容）
        /// </summary>
        public string Key
        {
            get => ReferenceMatchField;
            set => ReferenceMatchField = value;
        }

        /// <summary>
        /// 是否为必填字段（向后兼容）
        /// </summary>
        public bool Required
        {
            get => IsRequired;
            set => IsRequired = value;
        }

        /// <summary>
        /// 当字段为空时使用的默认值（向后兼容）
        /// </summary>
        public string DefaultValue
        {
            get => ImportNullValueReplacement?.ToString();
            set => ImportNullValueReplacement = value;
        }

        /// <summary>
        /// 显示名称（新架构兼容属性）
        /// </summary>
        public string DisplayName
        {
            get => ColumnName;
            set => ColumnName = value;
        }

        /// <summary>
        /// 是否可导入（新架构兼容属性）
        /// </summary>
        public bool CanImport
        {
            get => EnabledImport;
            set => EnabledImport = value;
        }

        /// <summary>
        /// 是否可导出（新架构兼容属性）
        /// </summary>
        public bool CanExport
        {
            get => EnabledExport;
            set => EnabledExport = value;
        }

        /// <summary>
        /// 排序顺序（向后兼容）
        /// </summary>
        public int Order
        {
            get => SortOrder;
            set => SortOrder = value;
        }

        #endregion

        #region 构造函数

        /// <summary>
        /// 默认构造函数
        /// </summary>
        public ExcelFieldAttribute()
        {
        }

        /// <summary>
        /// 构造函数，指定是否参与导入和导出
        /// </summary>
        /// <param name="enabledImport">是否参与导入</param>
        /// <param name="enabledExport">是否参与导出</param>
        public ExcelFieldAttribute(bool enabledImport = true, bool enabledExport = true)
        {
            EnabledImport = enabledImport;
            EnabledExport = enabledExport;
        }

        /// <summary>
        /// 构造函数，指定列名
        /// </summary>
        /// <param name="columnName">列名</param>
        public ExcelFieldAttribute(string columnName)
        {
            ColumnName = columnName;
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 获取有效的列名（优先使用ColumnName，否则使用属性名）
        /// </summary>
        /// <param name="propertyName">属性名称</param>
        /// <returns>有效的列名</returns>
        public string GetEffectiveColumnName(string propertyName)
        {
            return !string.IsNullOrEmpty(ColumnName) ? ColumnName : propertyName;
        }

        /// <summary>
        /// 检查是否有关联对象查找配置
        /// </summary>
        /// <returns>是否配置了ReferenceMatchField属性</returns>
        public bool HasKeyLookup()
        {
            return !string.IsNullOrEmpty(ReferenceMatchField) && ReferenceMatchField != "Oid";
        }

        /// <summary>
        /// 检查是否有默认值配置
        /// </summary>
        /// <returns>是否配置了默认值</returns>
        public bool HasDefaultValue()
        {
            return ImportNullValueReplacement != null;
        }

        /// <summary>
        /// 检查是否有自定义格式配置
        /// </summary>
        /// <returns>是否配置了自定义格式</returns>
        public bool HasCustomFormat()
        {
            return !string.IsNullOrEmpty(Format);
        }

        /// <summary>
        /// 检查是否有验证规则
        /// </summary>
        /// <returns>是否配置了验证规则</returns>
        public bool HasValidationRules()
        {
            return IsRequired || !string.IsNullOrEmpty(ValidationRegex) || !string.IsNullOrEmpty(Format);
        }

        /// <summary>
        /// 检查是否有自定义转换方法
        /// </summary>
        /// <returns>是否配置了转换方法</returns>
        public bool HasCustomConverter()
        {
            return !string.IsNullOrEmpty(ImportConvertMethod) || !string.IsNullOrEmpty(ExportConvertMethod);
        }

        /// <summary>
        /// 获取有效的验证错误消息
        /// </summary>
        /// <param name="propertyName">属性名称</param>
        /// <returns>验证错误消息</returns>
        public string GetEffectiveValidationErrorMessage(string propertyName)
        {
            if (!string.IsNullOrEmpty(ValidationErrorMessage))
            {
                return ValidationErrorMessage.Replace("{ColumnName}", GetEffectiveColumnName(propertyName));
            }
            return $"【{GetEffectiveColumnName(propertyName)}】格式错误";
        }

        /// <summary>
        /// 验证配置的有效性
        /// </summary>
        /// <returns>验证结果</returns>
        public (bool IsValid, string ErrorMessage) ValidateConfiguration()
        {
            if (ColumnIndex < 0)
            {
                return (false, "ColumnIndex不能小于0");
            }

            if (!string.IsNullOrEmpty(ValidationRegex))
            {
                try
                {
                    var regex = new System.Text.RegularExpressions.Regex(ValidationRegex);
                }
                catch (Exception ex)
                {
                    return (false, $"ValidationRegex格式错误: {ex.Message}");
                }
            }

            // 允许完全禁用集合属性、计算属性等不需要导入导出的字段
            // 如果没有启用任何功能，允许通过验证（用于明确禁用的字段）
            // 这种情况下，即使有其他配置也是合理的，例如标记为 Hidden 的集合属性

            return (true, null);
        }

        /// <summary>
        /// 返回特性的字符串表示
        /// </summary>
        /// <returns>特性描述</returns>
        public override string ToString()
        {
            var parts = new List<string>();
            
            if (!EnabledImport) parts.Add("EnabledImport=false");
            if (!EnabledExport) parts.Add("EnabledExport=false");
            if (!string.IsNullOrEmpty(ColumnName)) parts.Add($"ColumnName=\"{ColumnName}\"");
            if (ColumnIndex != 0) parts.Add($"ColumnIndex={ColumnIndex}");
            if (SortOrder != 0) parts.Add($"SortOrder={SortOrder}");
            if (IsRequired) parts.Add("IsRequired=true");
            if (!string.IsNullOrEmpty(Format)) parts.Add($"Format=\"{Format}\"");
            if (!string.IsNullOrEmpty(ValidationRegex)) parts.Add($"ValidationRegex=\"{ValidationRegex}\"");
            if (ImportNullValueReplacement != null) parts.Add($"ImportNullValueReplacement=\"{ImportNullValueReplacement}\"");
            if (!string.IsNullOrEmpty(ExportNullValueDisplay)) parts.Add($"ExportNullValueDisplay=\"{ExportNullValueDisplay}\"");
            if (Alignment != ExcelAlignment.Left) parts.Add($"Alignment={Alignment}");
            if (ColumnWidth != 0) parts.Add($"ColumnWidth={ColumnWidth}");
            if (IsBold) parts.Add("IsBold=true");
            if (Hidden) parts.Add("Hidden=true");
            if (!string.IsNullOrEmpty(ImportConvertMethod)) parts.Add($"ImportConvertMethod=\"{ImportConvertMethod}\"");
            if (!string.IsNullOrEmpty(ExportConvertMethod)) parts.Add($"ExportConvertMethod=\"{ExportConvertMethod}\"");
            if (ReferenceMatchField != "Oid") parts.Add($"ReferenceMatchField=\"{ReferenceMatchField}\"");
            if (ReferenceCreateIfNotExists) parts.Add("ReferenceCreateIfNotExists=true");

            return $"[ExcelField({string.Join(", ", parts)})]";
        }

        #endregion
    }
}