using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Actions;
using DevExpress.ExpressApp.SystemModule;
using DevExpress.Persistent.Base;
using DevExpress.Persistent.BaseImpl;
using Wxy.Xaf.DataDictionary;
using System.Linq;

namespace Wxy.Xaf.DataDictionary.Controllers;

/// <summary>
/// 数据字典列表视图控制器
/// </summary>
public class DataDictionaryListViewController : ViewController
{
    public DataDictionaryListViewController()
    {
        TargetViewType = ViewType.ListView;
        TargetObjectType = typeof(DataDictionary);
    }

    public SimpleAction ViewItemsAction { get; private set; }
    public SimpleAction AddNewItemAction { get; private set; }

    protected override void OnActivated()
    {
        base.OnActivated();

        // 查看字典项
        ViewItemsAction = new SimpleAction(this, "ViewItems", PredefinedCategory.View)
        {
            Caption = "查看字典项",
            ToolTip = "查看选中数据字典的所有字典项",
            ImageName = "BO_List",
            SelectionDependencyType = SelectionDependencyType.RequireSingleObject
        };
        ViewItemsAction.Execute += ViewItemsAction_Execute;

        // 添加新字典项
        AddNewItemAction = new SimpleAction(this, "AddNewItem", PredefinedCategory.Edit)
        {
            Caption = "添加字典项",
            ToolTip = "为选中的数据字典添加新的字典项",
            ImageName = "Action_New",
            SelectionDependencyType = SelectionDependencyType.RequireSingleObject
        };
        AddNewItemAction.Execute += AddNewItemAction_Execute;

        // 显示字典统计
        var showStatsAction = new SimpleAction(this, "ShowStatistics", "统计")
        {
            Caption = "统计信息",
            ToolTip = "显示选中数据字典的统计信息",
            ImageName = "BO_Summary",
            SelectionDependencyType = SelectionDependencyType.RequireSingleObject
        };
        showStatsAction.Execute += ShowStatsAction_Execute;
    }

    protected override void OnDeactivated()
    {
        if (ViewItemsAction != null)
        {
            ViewItemsAction.Execute -= ViewItemsAction_Execute;
        }
        if (AddNewItemAction != null)
        {
            AddNewItemAction.Execute -= AddNewItemAction_Execute;
        }
        base.OnDeactivated();
    }

    /// <summary>
    /// 查看字典项
    /// </summary>
    private void ViewItemsAction_Execute(object sender, SimpleActionExecuteEventArgs e)
    {
        var selectedDictionary = View.CurrentObject as DataDictionary;
        if (selectedDictionary == null) return;

        // 显示字典项列表
        var objectSpace = Application.CreateObjectSpace(typeof(DataDictionaryItem));
        var items = objectSpace.GetObjects<DataDictionaryItem>()
            .Where(item => item.DataDictionary?.Oid == selectedDictionary.Oid)
            .ToList();

        var message = $"数据字典: {selectedDictionary.Name}\n" +
                     $"字典项数量: {items.Count}\n\n";

        if (items.Any())
        {
            message += "字典项列表:\n";
            foreach (var item in items.OrderBy(i => i.Order))
            {
                message += $"  {item.Order}. {item.Name}";
                if (!string.IsNullOrEmpty(item.Code))
                {
                    message += $" ({item.Code})";
                }
                if (!string.IsNullOrEmpty(item.Description))
                {
                    message += $" - {item.Description}";
                }
                message += "\n";
            }
        }
        else
        {
            message += "暂无字典项";
        }

        Application.ShowViewStrategy.ShowMessage(message, InformationType.Info);
    }

    /// <summary>
    /// 添加新字典项
    /// </summary>
    private void AddNewItemAction_Execute(object sender, SimpleActionExecuteEventArgs e)
    {
        var selectedDictionary = View.CurrentObject as DataDictionary;
        if (selectedDictionary == null) return;

        // 使用新建对象控制器
        var newObjectViewController = Frame.GetController<NewObjectViewController>();
        if (newObjectViewController != null)
        {
            var firstItem = newObjectViewController.NewObjectAction.Items.FirstOrDefault();
            if (firstItem != null)
            {
                newObjectViewController.NewObjectAction.DoExecute(firstItem);
            }
        }
    }

    /// <summary>
    /// 显示统计信息
    /// </summary>
    private void ShowStatsAction_Execute(object sender, SimpleActionExecuteEventArgs e)
    {
        var selectedDictionary = View.CurrentObject as DataDictionary;
        if (selectedDictionary == null) return;

        var items = selectedDictionary.Items;

        var stats = $"数据字典统计信息\n\n" +
                   $"字典名称: {selectedDictionary.Name}\n" +
                   $"字典项总数: {items.Count}\n" +
                   $"最大顺序号: {(items.Any() ? items.Max(i => i.Order) : 0)}\n" +
                   $"有编码的项: {items.Count(i => !string.IsNullOrEmpty(i.Code))}\n" +
                   $"有描述的项: {items.Count(i => !string.IsNullOrEmpty(i.Description))}\n";

        if (items.Any())
        {
            var codes = items.Where(i => !string.IsNullOrEmpty(i.Code)).Select(i => i.Code).ToList();
            if (codes.Any())
            {
                stats += $"\n编码列表: {string.Join(", ", codes)}\n";
            }
        }

        Application.ShowViewStrategy.ShowMessage(stats, InformationType.Info);
    }
}
