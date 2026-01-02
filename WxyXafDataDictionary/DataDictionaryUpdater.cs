using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Updating;
using DevExpress.Xpo;
using DevExpress.Data.Filtering;
using System.Diagnostics;

namespace Wxy.Xaf.DataDictionary;

public class DataDictionaryUpdater : ModuleUpdater
{
    public DataDictionaryUpdater(IObjectSpace objectSpace, Version currentDBVersion)
        : base(objectSpace, currentDBVersion)
    {
    }

    public override void UpdateDatabaseAfterUpdateSchema()
    {
        base.UpdateDatabaseAfterUpdateSchema();

        // 每次都检查是否需要初始化数据字典
        // 检查数据库中是否已经有数据字典,如果没有则初始化
        try
        {
            var existingDictionaries = ObjectSpace.GetObjects<DataDictionary>();
            var hasExistingDictionaries = existingDictionaries != null && existingDictionaries.Count > 0;

            if (!hasExistingDictionaries)
            {
                Debug.WriteLine("=== DataDictionaryUpdater 数据库中没有数据字典,开始初始化 ===");
                InitializeDataDictionaries();
            }
            else
            {
                Debug.WriteLine($"=== DataDictionaryUpdater 数据库中已存在 {hasExistingDictionaries} 个数据字典,跳过初始化 ===");
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[DataDictionary] 检查数据字典时出错: {ex.Message}");
            // 如果检查失败,仍然尝试初始化
            InitializeDataDictionaries();
        }
    }

    private void InitializeDataDictionaries()
    {
        Debug.WriteLine("=== DataDictionaryUpdater 开始扫描数据字典 ===");

        try
        {
            // 收集所有使用 DataDictionaryAttribute 标记的属性信息
            var typesInfo = ObjectSpace.TypesInfo;
            var dictionaryConfigs = new Dictionary<string, (List<string> defaultItems, string description)>();

            foreach (var persistentTypeInfo in typesInfo.PersistentTypes)
            {
                var members = persistentTypeInfo.Members
                    .Where(m => m.IsProperty)
                    .Where(m => m.MemberType == typeof(XPCollection<DataDictionaryItem>) || m.MemberType == typeof(DataDictionaryItem));

                foreach (var member in members)
                {
                    var attribute = member.FindAttribute<DataDictionaryAttribute>();
                    if (attribute != null && !string.IsNullOrWhiteSpace(attribute.DataDictionaryName))
                    {
                        var dictName = attribute.DataDictionaryName;

                        if (!dictionaryConfigs.ContainsKey(dictName))
                        {
                            dictionaryConfigs[dictName] = (new List<string>(), attribute.Description ?? "");
                        }

                        if (!string.IsNullOrEmpty(attribute.DefaultItems))
                        {
                            var items = ParseDefaultItems(attribute.DefaultItems);
                            foreach (var item in items)
                            {
                                if (!dictionaryConfigs[dictName].defaultItems.Contains(item))
                                {
                                    dictionaryConfigs[dictName].defaultItems.Add(item);
                                }
                            }
                        }
                    }
                }
            }

            // 尝试使用 XPO 直接创建对象以绕过安全检查
            Session xpoSession = null;
            bool useXpoDirectly = false;
            try
            {
                var sessionProperty = ObjectSpace.GetType().GetProperty("Session",
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                xpoSession = sessionProperty?.GetValue(ObjectSpace) as Session;
                useXpoDirectly = xpoSession != null;
            }
            catch { }

            // 创建或更新数据字典及其默认项
            foreach (var config in dictionaryConfigs)
            {
                var dictName = config.Key;
                var defaultItems = config.Value.defaultItems;

                DataDictionary dataDictionary = null;

                if (useXpoDirectly)
                {
                    var existingDictsXpo = new XPCollection<DataDictionary>(xpoSession,
                        CriteriaOperator.Parse("Name = ?", dictName));
                    dataDictionary = existingDictsXpo.FirstOrDefault();

                    if (dataDictionary == null)
                    {
                        dataDictionary = new DataDictionary(xpoSession);
                        dataDictionary.Name = dictName;
                        dataDictionary.Save();
                    }
                }
                else
                {
                    dataDictionary = ObjectSpace.FirstOrDefault<DataDictionary>(d => d.Name == dictName);
                    if (dataDictionary == null)
                    {
                        dataDictionary = ObjectSpace.CreateObject<DataDictionary>();
                        dataDictionary.Name = dictName;
                    }
                }

                if (defaultItems.Any())
                {
                    InitializeDefaultItems(dataDictionary, defaultItems, useXpoDirectly ? xpoSession : null);
                }
            }

            // 提交更改
            if (useXpoDirectly)
            {
                if (xpoSession is DevExpress.Xpo.UnitOfWork uow)
                {
                    uow.CommitChanges();
                }
                Debug.WriteLine($"=== DataDictionaryUpdater 初始化完成 (创建 {dictionaryConfigs.Count} 个数据字典) ===");
            }
            else
            {
                ObjectSpace.CommitChanges();
                Debug.WriteLine($"=== DataDictionaryUpdater 初始化完成 (创建 {dictionaryConfigs.Count} 个数据字典) ===");
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[DataDictionary] 初始化失败: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// 解析默认项字符串
    /// 格式: "名称|编码;名称|编码" 或 "名称;名称"
    /// </summary>
    private List<string> ParseDefaultItems(string defaultItems)
    {
        var result = new List<string>();

        if (string.IsNullOrWhiteSpace(defaultItems))
            return result;

        // 按分号分割项
        var items = defaultItems.Split(new[] { ';', '；' }, StringSplitOptions.RemoveEmptyEntries);

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

    /// <summary>
    /// 初始化数据字典的默认项
    /// </summary>
    private void InitializeDefaultItems(DataDictionary dictionary, List<string> defaultItems, Session xpoSession = null)
    {
        int order = dictionary.Items.Count > 0 ? dictionary.Items.Max(i => i.Order) + 1 : 1;

        foreach (var itemConfig in defaultItems)
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
                DataDictionaryItem newItem;

                if (xpoSession != null)
                {
                    newItem = new DataDictionaryItem(xpoSession);
                }
                else
                {
                    newItem = ObjectSpace.CreateObject<DataDictionaryItem>();
                }

                newItem.Name = name;
                newItem.Code = code;
                newItem.Order = order++;
                newItem.DataDictionary = dictionary;

                if (xpoSession != null)
                {
                    newItem.Save();
                }
            }
        }
    }
}
