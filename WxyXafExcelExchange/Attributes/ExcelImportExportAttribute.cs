using System;
using Wxy.Xaf.ExcelExchange.Enums;

namespace Wxy.Xaf.ExcelExchange
{
    /// <summary>
    /// 类级别特性：控制整个XPO对象的Excel导入导出全局规则
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class ExcelImportExportAttribute : Attribute
    {
        #region 核心开关

        /// <summary>
        /// 是否启用该对象的Excel导入，默认为true
        /// </summary>
        public bool EnabledImport { get; set; } = true;

        /// <summary>
        /// 是否启用该对象的Excel导出，默认为true
        /// </summary>
        public bool EnabledExport { get; set; } = true;

        #endregion

        #region 工作表配置

        /// <summary>
        /// 指定导入/导出的Excel工作表名称
        /// 默认使用类的XAF DisplayName（本地化），如果没有则使用类名
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 文件名（SheetName的别名，提供更直观的参数名）
        /// </summary>
        public string FileName
        {
            get => SheetName;
            set => SheetName = value;
        }

        #endregion

        #region 导入配置

        /// <summary>
        /// 重复数据处理策略，默认为InsertOrUpdate
        /// </summary>
        public ImportDuplicateStrategy ImportDuplicateStrategy { get; set; } = ImportDuplicateStrategy.InsertOrUpdate;

        /// <summary>
        /// Excel表头行号（从1开始），默认为1
        /// </summary>
        public int HeaderRowIndex { get; set; } = 1;

        /// <summary>
        /// 数据起始行号，默认为2（表头下一行）
        /// </summary>
        public int DataStartRowIndex { get; set; } = 2;

        /// <summary>
        /// 导入完成后是否自动保存到数据库，默认为true
        /// </summary>
        public bool AutoSaveAfterImport { get; set; } = true;

        /// <summary>
        /// 导入验证模式，默认为Strict（严格模式）
        /// </summary>
        public ValidationMode ValidationMode { get; set; } = ValidationMode.Strict;

        #endregion

        #region 导出配置

        /// <summary>
        /// 导出时是否包含表头行，默认为true
        /// </summary>
        public bool ExportIncludeHeader { get; set; } = true;

        /// <summary>
        /// 导出时是否自动调整列宽，默认为true
        /// </summary>
        public bool ExportAutoAdjustColumnWidth { get; set; } = true;

        /// <summary>
        /// 导出模板文件名（如"CustomerTemplate.xlsx"）
        /// 用于基于模板的导出，支持样式和公式
        /// </summary>
        public string TemplateName { get; set; }

        /// <summary>
        /// 多对象批量导出时的排序优先级（数值越小越靠前），默认为0
        /// </summary>
        public int ExportSortOrder { get; set; } = 0;

        #endregion

        #region UI和描述

        /// <summary>
        /// UI显示的描述信息（如"客户信息导入导出"）
        /// 用于XAF界面的导入导出按钮提示
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// 导入操作的标题，默认为"从Excel导入"
        /// </summary>
        public string ImportCaption { get; set; } = "从Excel导入";

        /// <summary>
        /// 导出操作的标题，默认为"导出到Excel"
        /// </summary>
        public string ExportCaption { get; set; } = "导出到Excel";

        /// <summary>
        /// 默认文件名（不含扩展名），如果未指定则使用类名
        /// </summary>
        public string DefaultFileName { get; set; }

        #endregion

        #region 构造函数

        /// <summary>
        /// 默认构造函数
        /// </summary>
        public ExcelImportExportAttribute()
        {
        }

        /// <summary>
        /// 构造函数，指定工作表名称
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        public ExcelImportExportAttribute(string sheetName)
        {
            SheetName = sheetName;
        }

        /// <summary>
        /// 构造函数，指定导入导出开关
        /// </summary>
        /// <param name="enabledImport">是否启用导入</param>
        /// <param name="enabledExport">是否启用导出</param>
        public ExcelImportExportAttribute(bool enabledImport, bool enabledExport)
        {
            EnabledImport = enabledImport;
            EnabledExport = enabledExport;
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 获取有效的工作表名称
        /// 优先使用SheetName，否则尝试获取XAF DisplayName，最后使用类名
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <returns>有效的工作表名称</returns>
        public string GetEffectiveSheetName(Type objectType)
        {
            if (!string.IsNullOrEmpty(SheetName))
            {
                return SheetName;
            }

            // 尝试获取XAF DisplayName
            try
            {
                // 这里可以集成XAF的本地化机制
                // 暂时使用类名作为默认值
                return objectType.Name;
            }
            catch
            {
                return objectType.Name;
            }
        }

        /// <summary>
        /// 获取有效的默认文件名
        /// </summary>
        /// <param name="objectType">对象类型</param>
        /// <returns>有效的默认文件名</returns>
        public string GetEffectiveDefaultFileName(Type objectType)
        {
            return !string.IsNullOrEmpty(DefaultFileName) ? DefaultFileName : objectType.Name;
        }

        /// <summary>
        /// 验证配置的有效性
        /// </summary>
        /// <returns>验证结果</returns>
        public (bool IsValid, string ErrorMessage) ValidateConfiguration()
        {
            if (HeaderRowIndex < 1)
            {
                return (false, "HeaderRowIndex必须大于等于1");
            }

            if (DataStartRowIndex < 1)
            {
                return (false, "DataStartRowIndex必须大于等于1");
            }

            if (DataStartRowIndex <= HeaderRowIndex)
            {
                return (false, "DataStartRowIndex必须大于HeaderRowIndex");
            }

            if (!EnabledImport && !EnabledExport)
            {
                return (false, "至少需要启用导入或导出功能之一");
            }

            return (true, null);
        }

        /// <summary>
        /// 返回特性的字符串表示
        /// </summary>
        /// <returns>特性描述</returns>
        public override string ToString()
        {
            var parts = new System.Collections.Generic.List<string>();

            if (!EnabledImport) parts.Add("EnabledImport=false");
            if (!EnabledExport) parts.Add("EnabledExport=false");
            if (!string.IsNullOrEmpty(SheetName)) parts.Add($"SheetName=\"{SheetName}\"");
            if (ImportDuplicateStrategy != ImportDuplicateStrategy.InsertOrUpdate) 
                parts.Add($"ImportDuplicateStrategy={ImportDuplicateStrategy}");
            if (ValidationMode != ValidationMode.Strict) 
                parts.Add($"ValidationMode={ValidationMode}");
            if (HeaderRowIndex != 1) parts.Add($"HeaderRowIndex={HeaderRowIndex}");
            if (DataStartRowIndex != 2) parts.Add($"DataStartRowIndex={DataStartRowIndex}");
            if (!AutoSaveAfterImport) parts.Add("AutoSaveAfterImport=false");
            if (!ExportIncludeHeader) parts.Add("ExportIncludeHeader=false");
            if (!ExportAutoAdjustColumnWidth) parts.Add("ExportAutoAdjustColumnWidth=false");
            if (!string.IsNullOrEmpty(TemplateName)) parts.Add($"TemplateName=\"{TemplateName}\"");
            if (ExportSortOrder != 0) parts.Add($"ExportSortOrder={ExportSortOrder}");
            if (!string.IsNullOrEmpty(Description)) parts.Add($"Description=\"{Description}\"");

            return $"[ExcelImportExport({string.Join(", ", parts)})]";
        }

        #endregion
    }
}