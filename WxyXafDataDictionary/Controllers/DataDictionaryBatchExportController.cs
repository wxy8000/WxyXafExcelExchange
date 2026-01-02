using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Actions;
using DevExpress.Persistent.Base;
using Wxy.Xaf.ExcelExchange.Services;
using Wxy.Xaf.DataDictionary;
using Wxy.Xaf.DataDictionary;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;

namespace Wxy.Xaf.DataDictionary.Controllers;

/// <summary>
/// 数据字典批量导入导出控制器
/// 支持将多个字典导出到单个 Excel 文件(每个字典一个 Sheet)
/// 支持从多 Sheet Excel 文件导入字典数据
/// </summary>
public class DataDictionaryBatchExportController : ViewController
{
    public DataDictionaryBatchExportController()
    {
        // 只在数据字典列表视图中激活
        TargetViewType = ViewType.ListView;
        TargetObjectType = typeof(DataDictionary);
    }

    public SimpleAction ExportSelectedDictionariesAction { get; private set; }
    public SimpleAction ExportAllDictionariesAction { get; private set; }
    public SimpleAction ImportDictionariesAction { get; private set; }

    protected override void OnActivated()
    {
        base.OnActivated();

        // 导出选中的字典
        ExportSelectedDictionariesAction = new SimpleAction(this, "ExportSelectedDictionaries", PredefinedCategory.Export)
        {
            Caption = "导出选中的字典",
            ToolTip = "将选中的字典导出到单个 Excel 文件,每个字典为一个 Sheet",
            ImageName = "Export"
        };
        ExportSelectedDictionariesAction.Execute += ExportSelectedDictionariesAction_Execute;

        // 导出所有字典
        ExportAllDictionariesAction = new SimpleAction(this, "ExportAllDictionaries", PredefinedCategory.Export)
        {
            Caption = "导出所有字典",
            ToolTip = "将所有字典导出到单个 Excel 文件,每个字典为一个 Sheet",
            ImageName = "Export"
        };
        ExportAllDictionariesAction.Execute += ExportAllDictionariesAction_Execute;

        // 导入字典数据
        ImportDictionariesAction = new SimpleAction(this, "ImportDictionaries", "导入")
        {
            Caption = "导入字典数据",
            ToolTip = "从 Excel 文件导入字典数据,支持多 Sheet 导入",
            ImageName = "Import"
        };
        ImportDictionariesAction.Execute += ImportDictionariesAction_Execute;
    }

    protected override void OnDeactivated()
    {
        if (ExportSelectedDictionariesAction != null)
        {
            ExportSelectedDictionariesAction.Execute -= ExportSelectedDictionariesAction_Execute;
        }
        if (ExportAllDictionariesAction != null)
        {
            ExportAllDictionariesAction.Execute -= ExportAllDictionariesAction_Execute;
        }
        if (ImportDictionariesAction != null)
        {
            ImportDictionariesAction.Execute -= ImportDictionariesAction_Execute;
        }
        base.OnDeactivated();
    }

    /// <summary>
    /// 导出选中的字典
    /// </summary>
    private async void ExportSelectedDictionariesAction_Execute(object sender, SimpleActionExecuteEventArgs e)
    {
        var selectedObjects = View.SelectedObjects.Cast<DataDictionary>().ToList();
        if (selectedObjects.Count == 0)
        {
            Application.ShowViewStrategy.ShowMessage("请先选择要导出的字典", InformationType.Warning);
            return;
        }

        await ExportDictionariesAsync(selectedObjects, "选中的字典");
    }

    /// <summary>
    /// 导出所有字典
    /// </summary>
    private async void ExportAllDictionariesAction_Execute(object sender, SimpleActionExecuteEventArgs e)
    {
        var objectSpace = Application.CreateObjectSpace(typeof(DataDictionary));
        try
        {
            var allDictionaries = objectSpace.GetObjects<DataDictionary>().ToList();
            if (allDictionaries.Count == 0)
            {
                Application.ShowViewStrategy.ShowMessage("没有可导出的字典", InformationType.Warning);
                return;
            }

            await ExportDictionariesAsync(allDictionaries, "所有字典");
        }
        finally
        {
            objectSpace.Dispose();
        }
    }

    /// <summary>
    /// 导出字典到 Excel 文件
    /// </summary>
    private async Task ExportDictionariesAsync(List<DataDictionary> dictionaries, string description)
    {
        try
        {
            // 准备数据: 每个字典的所有字典项
            var exportData = new Dictionary<string, List<DataDictionaryItem>>();
            foreach (var dictionary in dictionaries)
            {
                // 使用当前视图的 ObjectSpace 获取字典项(确保加载完整数据)
                var items = ObjectSpace.GetObjects<DataDictionaryItem>()
                    .Where(item => item.DataDictionary?.Oid == dictionary.Oid)
                    .OrderBy(item => item.Order)
                    .ToList();

                // Sheet 名称使用字典名称(清理 Excel Sheet 名称中的非法字符)
                var sheetName = CleanSheetName(dictionary.Name);
                exportData[sheetName] = items;
            }

            // 获取 Excel 导入导出服务
            var exportService = ObjectSpace.ServiceProvider?.GetService(typeof(IExcelImportExportService)) as IExcelImportExportService;
            if (exportService == null)
            {
                Application.ShowViewStrategy.ShowMessage("Excel 导入导出服务不可用", InformationType.Error);
                return;
            }

            // 使用文件对话框选择保存路径
            var filePath = PromptForFileSave($"导出{description}_{System.DateTime.Now:yyyyMMddHHmmss}.xlsx");
            if (string.IsNullOrEmpty(filePath))
            {
                return; // 用户取消了操作
            }

            // 导出数据(每个字典一个 Sheet)
            await exportService.ExportMultipleSheetsAsync(
                filePath,
                exportData,
                new[] { "名称", "编码", "描述", "顺序" }, // 表头
                item => new object[] // 数据提取器
                {
                    item.Name,
                    item.Code ?? "",
                    item.Description ?? "",
                    item.Order
                },
                "数据字典导出"
            );

            Application.ShowViewStrategy.ShowMessage($"成功导出 {dictionaries.Count} 个字典到文件: {filePath}", InformationType.Success);
        }
        catch (System.Exception ex)
        {
            Application.ShowViewStrategy.ShowMessage($"导出失败: {ex.Message}", InformationType.Error);
        }
    }

    /// <summary>
    /// 导入字典数据
    /// </summary>
    private async void ImportDictionariesAction_Execute(object sender, SimpleActionExecuteEventArgs e)
    {
        try
        {
            // 使用文件对话框选择要导入的文件
            var filePath = PromptForFileOpen();
            if (string.IsNullOrEmpty(filePath))
            {
                return; // 用户取消了操作
            }

            // 获取 Excel 导入导出服务
            var importService = ObjectSpace.ServiceProvider?.GetService(typeof(IExcelImportExportService)) as IExcelImportExportService;
            if (importService == null)
            {
                Application.ShowViewStrategy.ShowMessage("Excel 导入导出服务不可用", InformationType.Error);
                return;
            }

            // 读取 Excel 文件的所有 Sheet
            var sheetsData = await importService.GetExcelSheetsAsync(filePath);
            if (sheetsData == null || sheetsData.Count == 0)
            {
                Application.ShowViewStrategy.ShowMessage("Excel 文件中没有检测到数据", InformationType.Warning);
                return;
            }

            // 处理每个 Sheet
            var importedCount = 0;
            var errorMessages = new List<string>();

            foreach (var sheetData in sheetsData)
            {
                var sheetName = sheetData.Key;
                var data = sheetData.Value;

                if (data == null || data.Count == 0)
                {
                    continue;
                }

                try
                {
                    // 查找或创建对应的字典
                    var dictionary = FindOrCreateDictionary(sheetName);
                    if (dictionary == null)
                    {
                        errorMessages.Add($"Sheet '{sheetName}': 无法创建或查找对应的字典");
                        continue;
                    }

                    // 导入字典项数据
                    var importResult = await ImportDictionaryItemsAsync(dictionary, data);
                    var success = importResult.Item1;
                    var errors = importResult.Item2;
                    importedCount += success;

                    if (errors.Any())
                    {
                        errorMessages.AddRange(errors.Select(err => $"Sheet '{sheetName}': {err}"));
                    }
                }
                catch (System.Exception ex)
                {
                    errorMessages.Add($"Sheet '{sheetName}': {ex.Message}");
                }
            }

            // 提交修改
            ObjectSpace.CommitChanges();

            // 显示导入结果
            var message = $"导入完成!\n成功导入 {importedCount} 条字典项";
            if (errorMessages.Any())
            {
                message += $"\n\n注意到 {errorMessages.Count} 个错误\n" + string.Join("\n", errorMessages.Take(10));
                if (errorMessages.Count > 10)
                {
                    message += $"\n... 还有 {errorMessages.Count - 10} 个错误";
                }
            }

            Application.ShowViewStrategy.ShowMessage(message, errorMessages.Any() ? InformationType.Warning : InformationType.Success);

            // 刷新视图
            View.Refresh();
        }
        catch (System.Exception ex)
        {
            Application.ShowViewStrategy.ShowMessage($"导入失败: {ex.Message}", InformationType.Error);
        }
    }

    /// <summary>
    /// 查找或创建字典
    /// </summary>
    private DataDictionary FindOrCreateDictionary(string dictionaryName)
    {
        // 先查找已存在的字典
        var existingDictionary = ObjectSpace.GetObjects<DataDictionary>()
            .FirstOrDefault(d => d.Name == dictionaryName);

        if (existingDictionary != null)
        {
            return existingDictionary;
        }

        // 创建新字典
        var newDictionary = ObjectSpace.CreateObject<DataDictionary>();
        newDictionary.Name = dictionaryName;
        return newDictionary;
    }

    /// <summary>
    /// 导入字典项数据
    /// </summary>
    private async Task<(int success, List<string> errors)> ImportDictionaryItemsAsync(
        DataDictionary dictionary,
        List<Dictionary<string, object>> sheetData)
    {
        var successCount = 0;
        var errors = new List<string>();

        // 跳过表头行(假设第一行是表头)
        var dataRows = sheetData.Skip(1).ToList();

        foreach (var row in dataRows)
        {
            try
            {
                // 提取数据(支持中英文列名)
                var name = GetStringValue(row, "名称") ?? GetStringValue(row, "Name");
                var code = GetStringValue(row, "编码") ?? GetStringValue(row, "Code");
                var description = GetStringValue(row, "描述") ?? GetStringValue(row, "Description");
                var order = GetIntValue(row, "顺序") ?? GetIntValue(row, "Order") ?? 0;

                // 验证必填字段
                if (string.IsNullOrWhiteSpace(name))
                {
                    errors.Add($"行 {dataRows.IndexOf(row) + 2}: 名称不能为空");
                    continue;
                }

                // 查找已存在的字典项
                var existingItem = dictionary.Items.FirstOrDefault(item => item.Name == name);

                if (existingItem != null)
                {
                    // 更新已存在的字典项
                    existingItem.Code = code;
                    existingItem.Description = description;
                    existingItem.Order = order;
                }
                else
                {
                    // 创建新字典项
                    var newItem = ObjectSpace.CreateObject<DataDictionaryItem>();
                    newItem.Name = name;
                    newItem.Code = code;
                    newItem.Description = description;
                    newItem.Order = order;
                    // DataDictionary 会在 OnSaving 中自动设置
                }

                successCount++;
            }
            catch (System.Exception ex)
            {
                errors.Add($"行 {dataRows.IndexOf(row) + 2}: {ex.Message}");
            }
        }

        return await Task.FromResult((successCount, errors));
    }

    /// <summary>
    /// 从字典中获取字符串值
    /// </summary>
    private string GetStringValue(Dictionary<string, object> data, string key)
    {
        if (data.TryGetValue(key, out var value) && value != null)
        {
            return value.ToString();
        }
        return null;
    }

    /// <summary>
    /// 从字典中获取整数值
    /// </summary>
    private int? GetIntValue(Dictionary<string, object> data, string key)
    {
        if (data.TryGetValue(key, out var value) && value != null)
        {
            if (int.TryParse(value.ToString(), out var result))
            {
                return result;
            }
        }
        return null;
    }

    /// <summary>
    /// 清理 Sheet 名称,移除 Excel 中不支持的字符
    /// </summary>
    private string CleanSheetName(string sheetName)
    {
        // Excel Sheet 名称不能包含的字符: \ / ? * [ ] :
        var invalidChars = new char[] { '\\', '/', '?', '*', '[', ']', ':' };
        var cleaned = sheetName;

        foreach (var ch in invalidChars)
        {
            cleaned = cleaned.Replace(ch, '_');
        }

        // Excel Sheet 名称最大长度为 31 个字符
        if (cleaned.Length > 31)
        {
            cleaned = cleaned.Substring(0, 31);
        }

        return cleaned;
    }

    /// <summary>
    /// 提示用户选择保存文件路径
    /// </summary>
    private string PromptForFileSave(string defaultFileName)
    {
        #if WinForms
            using (var saveDialog = new System.Windows.Forms.SaveFileDialog())
            {
                saveDialog.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
                saveDialog.FileName = defaultFileName;
                saveDialog.Title = "选择保存位置";

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return saveDialog.FileName;
                }
            }
        #endif

        return null;
    }

    /// <summary>
    /// 提示用户选择要打开的文件
    /// </summary>
    private string PromptForFileOpen()
    {
        #if WinForms
            using (var openDialog = new System.Windows.Forms.OpenFileDialog())
            {
                openDialog.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
                openDialog.Title = "选择要导入的 Excel 文件";
                openDialog.Multiselect = false;

                if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return openDialog.FileName;
                }
            }
        #endif

        return null;
    }
}
