using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Blazor;
using DevExpress.ExpressApp.DC;
using DevExpress.ExpressApp.Updating;
using DevExpress.Persistent.Base;
using DevExpress.Xpo;
using Microsoft.JSInterop;
using System.ComponentModel;

namespace Wxy.Xaf.DataDictionary;

public sealed class WxyXafDataDictionaryModule : ModuleBase
{
    public WxyXafDataDictionaryModule()
    {
        RequiredModuleTypes.Add(typeof(DevExpress.ExpressApp.SystemModule.SystemModule));
    }

    public override IEnumerable<ModuleUpdater> GetModuleUpdaters(IObjectSpace objectSpace, Version versionFromDB)
    {
        return new ModuleUpdater[]
        {
            new DataDictionaryUpdater(objectSpace, versionFromDB)
        };
    }

    public override void Setup(XafApplication application)
    {
        base.Setup(application);

        // 在 Blazor 应用中注入 CSS 以禁用数据字典链接
        if (application is BlazorApplication blazorApp)
        {
            InjectDisableLinksCss(blazorApp);
        }
    }

    private void InjectDisableLinksCss(BlazorApplication blazorApp)
    {
        // 使用多次尝试确保CSS注入成功
        System.Threading.Tasks.Task.Delay(1000).ContinueWith(_ =>
        {
            InjectCss(blazorApp);
        });

        // 再次尝试,确保在导航完全加载后执行
        System.Threading.Tasks.Task.Delay(3000).ContinueWith(_ =>
        {
            InjectCss(blazorApp);
        });

        // 第三次尝试,防止页面刷新后失效
        System.Threading.Tasks.Task.Delay(5000).ContinueWith(_ =>
        {
            InjectCss(blazorApp);
        });
    }

    private void InjectCss(BlazorApplication blazorApp)
    {
        try
        {
            var jsRuntime = blazorApp.ServiceProvider?.GetService(typeof(IJSRuntime)) as IJSRuntime;
            if (jsRuntime != null)
            {
                // 只在列表视图的表格单元格内禁用数据字典链接,不影响左侧导航和其他地方
                string js = @"(function() {
                    var styleId = 'data-dictionary-disable-links';
                    var existingStyle = document.getElementById(styleId);
                    if (existingStyle) existingStyle.remove();

                    var style = document.createElement('style');
                    style.id = styleId;
                    // 精确定位:只在数据表格的tbody内的链接,排除导航菜单
                    style.textContent = `
                        /* 主内容区域的表格 */
                        .main-content table a[href*='DataDictionaryItem'],
                        .main-content table a[href*='DataDictionary'],

                        /* 列表视图的表格 */
                        .list-view table a[href*='DataDictionaryItem'],
                        .list-view table a[href*='DataDictionary'],

                        /* 表格行内的链接 */
                        tbody tr a[href*='DataDictionaryItem'],
                        tbody tr a[href*='DataDictionary'],

                        /* DevExpress Blazor 网格组件 */
                        .dx-datagrid tbody a[href*='DataDictionaryItem'],
                        .dx-datagrid tbody a[href*='DataDictionary']

                        {
                            pointer-events: none !important;
                            color: inherit !important;
                            text-decoration: none !important;
                            cursor: text !important;
                        }

                        /* 明确排除导航菜单 */
                        .nav-menu a[href*='DataDictionaryItem'],
                        .nav-menu a[href*='DataDictionary'],
                        .sidebar a[href*='DataDictionaryItem'],
                        .sidebar a[href*='DataDictionary'],
                        .navigation a[href*='DataDictionaryItem'],
                        .navigation a[href*='DataDictionary'],
                        aside a[href*='DataDictionaryItem'],
                        aside a[href*='DataDictionary']
                        {
                            pointer-events: auto !important;
                            cursor: pointer !important;
                        }
                    `;
                    document.head.appendChild(style);

                    console.log('[DataDictionary] CSS injected for table cells only (excluding navigation)');

                    // 统计表格内的链接
                    var tableLinks = document.querySelectorAll('.main-content table a[href*=\'DataDictionary\'], .list-view table a[href*=\'DataDictionary\']');
                    console.log('[DataDictionary] Found ' + tableLinks.length + ' DataDictionary links in table views');
                })();";

                System.Threading.Tasks.Task.Run(async () =>
                {
                    try
                    {
                        await jsRuntime.InvokeVoidAsync("eval", js);
                    }
                    catch
                    {
                        // 静默处理JS调用错误
                    }
                });
            }
        }
        catch
        {
            // 静默处理错误，避免影响应用启动
        }
    }

    public override void CustomizeTypesInfo(ITypesInfo typesInfo)
    {
        base.CustomizeTypesInfo(typesInfo);

        // 扫描所有持久化类型，查找使用DataDictionaryAttribute的属性
        foreach (var persistentTypeInfo in typesInfo.PersistentTypes.Where(p => p.IsPersistent))
        {
            // 查找DataDictionaryItem类型的属性成员
            var members = persistentTypeInfo.Members
                .Where(m => m.IsProperty)
                .Where(m => m.MemberType == typeof(XPCollection<DataDictionaryItem>)
                        || m.MemberType == typeof(DataDictionaryItem));

            foreach (var member in members)
            {
                // 查找DataDictionaryAttribute特性
                var attribute = member.FindAttribute<DataDictionaryAttribute>();
                if (attribute != null && !string.IsNullOrWhiteSpace(attribute.DataDictionaryName))
                {
                    // 获取DataDictionaryItem的TypeInfo
                    var dictItemTypeInfo = typesInfo.FindTypeInfo(typeof(DataDictionaryItem));

                    // 在DataDictionaryItem中添加反向关联成员
                    var dictItemMember = dictItemTypeInfo.CreateMember(
                        $"{persistentTypeInfo.Name}_{member.Name}",
                        typeof(XPCollection<>).MakeGenericType(persistentTypeInfo.Type));

                    // 创建唯一的关联名称
                    var associationName = $"{persistentTypeInfo.Name}_{member.Name}_{nameof(DataDictionaryItem)}";

                    // 添加双向关联属性
                    dictItemMember.AddAttribute(new AssociationAttribute(associationName, persistentTypeInfo.Type), true);
                    dictItemMember.AddAttribute(new BrowsableAttribute(false), true);
                    member.AddAttribute(new AssociationAttribute(associationName, typeof(DataDictionaryItem)), true);

                    // 添加数据源过滤条件（如果不存在）
                    if (member.FindAttribute<DataSourceCriteriaAttribute>() == null
                        && member.FindAttribute<DataSourceCriteriaPropertyAttribute>() == null
                        && member.FindAttribute<DataSourcePropertyAttribute>() == null)
                    {
                        var criteria = $"[DataDictionary.Name]='{attribute.DataDictionaryName}'";
                        member.AddAttribute(new DataSourceCriteriaAttribute(criteria), true);
                    }

                    // 刷新成员信息
                    ((XafMemberInfo)member).Refresh();
                    ((XafMemberInfo)dictItemMember).Refresh();
                }
            }
        }
    }
}
