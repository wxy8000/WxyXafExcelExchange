using DevExpress.Persistent.Base;
using DevExpress.Persistent.BaseImpl;
using DevExpress.Persistent.Validation;
using DevExpress.Xpo;
using Wxy.Xaf.ExcelExchange;
using Wxy.Xaf.DataDictionary;

namespace Wxy.Xaf.DataDictionary;

[Persistent]
[NavigationItem]
[System.ComponentModel.DisplayName("数据字典")]
[ExcelImportExport(EnabledImport = true, EnabledExport = true, DataStartRowIndex = 2)]
public class DataDictionary : BaseObject
{
    private string _name;

    [RuleRequiredField]
    [RuleUniqueValue]
    [System.ComponentModel.DisplayName("名称")]
    [ExcelField(SortOrder = 1, ColumnName = "字典名称")]
    public string Name
    {
        get => _name;
        set => SetPropertyValue(nameof(Name), ref _name, value);
    }

    [System.ComponentModel.DisplayName("字典项")]
    [Association("DataDictionary-Items"), Aggregated]
    public XPCollection<DataDictionaryItem> Items
    {
        get => GetCollection<DataDictionaryItem>(nameof(Items));
    }

    public DataDictionary(Session session)
        : base(session)
    {
    }
}

