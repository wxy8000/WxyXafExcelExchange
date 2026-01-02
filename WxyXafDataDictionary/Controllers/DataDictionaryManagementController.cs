using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Actions;
using DevExpress.Persistent.Base;
using System.Diagnostics;
using System.Linq;
using Wxy.Xaf.DataDictionary;

namespace Wxy.Xaf.DataDictionary.Controllers;

/// <summary>
/// 数据字典管理控制器
/// 提供数据字典的初始化、刷新和管理功能
/// </summary>
public class DataDictionaryManagementController : ViewController
{
    public DataDictionaryManagementController()
    {
        // 在数据字典列表视图中激活
        TargetViewType = ViewType.ListView;
        TargetObjectType = typeof(DataDictionary);
    }

    public SimpleAction RefreshDictionariesAction { get; private set; }
    public SimpleAction ScanDictionaryPropertiesAction { get; private set; }

    protected override void OnActivated()
    {
        base.OnActivated();

        // 刷新数据字典
        RefreshDictionariesAction = new SimpleAction(this, "RefreshDictionaries", "管理")
        {
            Caption = "刷新数据字典",
            ToolTip = "扫描所有 DataDictionaryAttribute 标记的属性,自动创建缺失的数据字典",
            ImageName = "Action_Refresh"
        };
        RefreshDictionariesAction.Execute += RefreshDictionariesAction_Execute;

        // 扫描字典属性
        ScanDictionaryPropertiesAction = new SimpleAction(this, "ScanDictionaryProperties", "管理")
        {
            Caption = "扫描字典属性",
            ToolTip = "扫描系统中所有使用数据字典的属性,并显示统计信息",
            ImageName = "BO_Scan"
        };
        ScanDictionaryPropertiesAction.Execute += ScanDictionaryPropertiesAction_Execute;
    }

    protected override void OnDeactivated()
    {
        if (RefreshDictionariesAction != null)
        {
            RefreshDictionariesAction.Execute -= RefreshDictionariesAction_Execute;
        }
        if (ScanDictionaryPropertiesAction != null)
        {
            ScanDictionaryPropertiesAction.Execute -= ScanDictionaryPropertiesAction_Execute;
        }
        base.OnDeactivated();
    }

    /// <summary>
    /// 刷新数据字典
    /// </summary>
    private void RefreshDictionariesAction_Execute(object sender, SimpleActionExecuteEventArgs e)
    {
        try
        {
            Debug.WriteLine("=== 开始刷新数据字典 ===");

            var typesInfo = Application.TypesInfo;
            var createdCount = 0;
            var updatedCount = 0;

            foreach (var persistentTypeInfo in typesInfo.PersistentTypes)
            {
                var members = persistentTypeInfo.Members
                    .Where(m => m.IsProperty)
                    .Where(m => m.MemberType.Name.Contains("DataDictionaryItem"));

                foreach (var member in members)
                {
                    var attribute = member.FindAttribute<DataDictionaryAttribute>();
                    if (attribute != null && !string.IsNullOrWhiteSpace(attribute.DataDictionaryName))
                    {
                        var dictName = attribute.DataDictionaryName;

                        // 查找或创建数据字典
                        var dataDictionary = ObjectSpace.FirstOrDefault<DataDictionary>(d => d.Name == dictName);
                        if (dataDictionary == null)
                        {
                            dataDictionary = ObjectSpace.CreateObject<DataDictionary>();
                            dataDictionary.Name = dictName;
                            createdCount++;
                            Debug.WriteLine($"  创建数据字典: {dictName}");
                        }
                        else
                        {
                            updatedCount++;
                            Debug.WriteLine($"  已存在数据字典: {dictName}");
                        }

                        // 如果有默认项,初始化它们
                        if (!string.IsNullOrEmpty(attribute.DefaultItems))
                        {
                            InitializeDefaultItemsForDictionary(dataDictionary, attribute.DefaultItems);
                        }
                    }
                }
            }

            ObjectSpace.CommitChanges();

            var message = $"数据字典刷新完成!\n新建: {createdCount} 个\n已存在: {updatedCount} 个";
            Application.ShowViewStrategy.ShowMessage(message, InformationType.Success);

            Debug.WriteLine("=== 数据字典刷新完成 ===");

            // 刷新视图
            View.Refresh();
        }
        catch (System.Exception ex)
        {
            Application.ShowViewStrategy.ShowMessage($"刷新失败: {ex.Message}", InformationType.Error);
            Debug.WriteLine($"刷新失败: {ex.Message}\n{ex.StackTrace}");
        }
    }

    /// <summary>
    /// 扫描字典属性
    /// </summary>
    private void ScanDictionaryPropertiesAction_Execute(object sender, SimpleActionExecuteEventArgs e)
    {
        try
        {
            Debug.WriteLine("=== 开始扫描数据字典属性 ===");

            var typesInfo = Application.TypesInfo;
            var propertyDict = new System.Collections.Generic.SortedDictionary<string, System.Collections.Generic.List<string>>();

            foreach (var persistentTypeInfo in typesInfo.PersistentTypes)
            {
                var members = persistentTypeInfo.Members
                    .Where(m => m.IsProperty)
                    .Where(m => m.MemberType.Name.Contains("DataDictionaryItem"));

                foreach (var member in members)
                {
                    var attribute = member.FindAttribute<DataDictionaryAttribute>();
                    if (attribute != null && !string.IsNullOrWhiteSpace(attribute.DataDictionaryName))
                    {
                        var dictName = attribute.DataDictionaryName;
                        var propertyInfo = $"{persistentTypeInfo.Type.Name}.{member.Name}";

                        if (!propertyDict.ContainsKey(dictName))
                        {
                            propertyDict[dictName] = new System.Collections.Generic.List<string>();
                        }
                        propertyDict[dictName].Add(propertyInfo);

                        Debug.WriteLine($"  找到: {propertyInfo} -> {dictName}");
                    }
                }
            }

            // 生成统计报告
            var report = new System.Text.StringBuilder();
            report.AppendLine($"数据字典扫描结果");
            report.AppendLine($"共找到 {propertyDict.Count} 个数据字典:");
            report.AppendLine();

            foreach (var kvp in propertyDict)
            {
                report.AppendLine($"字典: {kvp.Key}");
                report.AppendLine($"  使用属性数: {kvp.Value.Count}");
                foreach (var prop in kvp.Value)
                {
                    report.AppendLine($"    - {prop}");
                }
                report.AppendLine();
            }

            Debug.WriteLine("=== 扫描完成 ===");
            Debug.WriteLine(report.ToString());

            Application.ShowViewStrategy.ShowMessage(report.ToString(), InformationType.Info);
        }
        catch (System.Exception ex)
        {
            Application.ShowViewStrategy.ShowMessage($"扫描失败: {ex.Message}", InformationType.Error);
            Debug.WriteLine($"扫描失败: {ex.Message}\n{ex.StackTrace}");
        }
    }

    /// <summary>
    /// 为字典初始化默认项
    /// </summary>
    private void InitializeDefaultItemsForDictionary(DataDictionary dictionary, string defaultItems)
    {
        var items = ParseDefaultItems(defaultItems);
        int order = 1;

        foreach (var itemConfig in items)
        {
            string name, code = "";

            if (itemConfig.Contains("|"))
            {
                var parts = itemConfig.Split('|');
                name = parts[0].Trim();
                code = parts.Length > 1 ? parts[1].Trim() : "";
            }
            else
            {
                name = itemConfig;
                code = "";
            }

            var existingItem = dictionary.Items.FirstOrDefault(i => i.Name == name);
            if (existingItem == null)
            {
                var newItem = ObjectSpace.CreateObject<DataDictionaryItem>();
                newItem.Name = name;
                newItem.Code = code;
                newItem.Order = order++;
                newItem.DataDictionary = dictionary;

                Debug.WriteLine($"    创建默认项: {name} ({code})");
            }
        }
    }

    /// <summary>
    /// 解析默认项字符串
    /// </summary>
    private System.Collections.Generic.List<string> ParseDefaultItems(string defaultItems)
    {
        var result = new System.Collections.Generic.List<string>();

        if (string.IsNullOrWhiteSpace(defaultItems))
            return result;

        var items = defaultItems.Split(new[] { ';', '；' }, System.StringSplitOptions.RemoveEmptyEntries);

        foreach (var item in items)
        {
            var trimmed = item.Trim();
            if (!string.IsNullOrEmpty(trimmed))
            {
                result.Add(trimmed);
            }
        }

        return result;
    }
}
