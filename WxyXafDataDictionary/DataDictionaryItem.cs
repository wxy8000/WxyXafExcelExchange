using DevExpress.Persistent.BaseImpl;
using DevExpress.Persistent.Validation;
using DevExpress.Xpo;
using System.ComponentModel;
using Wxy.Xaf.ExcelExchange;
using Wxy.Xaf.DataDictionary;

namespace Wxy.Xaf.DataDictionary;

[Persistent]
[System.ComponentModel.DisplayName("字典项")]
[ExcelImportExport(EnabledImport = true, EnabledExport = true, DataStartRowIndex = 2)]
[DefaultProperty("Name")]
public class DataDictionaryItem : BaseObject
{
    private string _name;
    private string _code;
    private string _description;
    private int _order;
    private DataDictionary _dataDictionary;

    [RuleRequiredField]
    [System.ComponentModel.DisplayName("名称")]
    [ExcelField(SortOrder = 1, ColumnName = "名称", IsRequired = true)]
    public string Name
    {
        get => _name;
        set => SetPropertyValue(nameof(Name), ref _name, value);
    }

    [System.ComponentModel.DisplayName("编码")]
    [ExcelField(SortOrder = 2, ColumnName = "编码")]
    public string Code
    {
        get => _code;
        set => SetPropertyValue(nameof(Code), ref _code, value);
    }

    [System.ComponentModel.DisplayName("描述")]
    [ExcelField(SortOrder = 3, ColumnName = "描述")]
    public string Description
    {
        get => _description;
        set => SetPropertyValue(nameof(Description), ref _description, value);
    }

    [System.ComponentModel.DisplayName("顺序")]
    [ExcelField(SortOrder = 4, ColumnName = "顺序")]
    public int Order
    {
        get => _order;
        set => SetPropertyValue(nameof(Order), ref _order, value);
    }

    [Association("DataDictionary-Items")]
    [ExcelField(Hidden = true)] // 关联字段,在导入时通过 Sheet 名称处理
    public DataDictionary DataDictionary
    {
        get => _dataDictionary;
        set => SetPropertyValue(nameof(DataDictionary), ref _dataDictionary, value);
    }

    [Browsable(false)]
    [ExcelField(Hidden = true)]
    [RuleFromBoolProperty("字典项名称必须唯一", DefaultContexts.Save, CustomMessageTemplate = "字典项名称已存在")]
    public bool IsNameUnique
    {
        get => DataDictionary == null || !DataDictionary.Items.Any(item => item.Name == Name && item != this);
    }

    public DataDictionaryItem(Session session)
        : base(session)
    {
    }

    protected override void OnSaving()
    {
        base.OnSaving();

        // 自动设置Order：如果未设置且有父字典，自动设置为最大顺序号+1
        if (DataDictionary != null && Order == 0 && DataDictionary.Items.Count > 0)
        {
            Order = DataDictionary.Items.Max(item => item.Order) + 1;
        }
    }
}
