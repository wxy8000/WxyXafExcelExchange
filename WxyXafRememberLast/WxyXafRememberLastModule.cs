using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Model;
using Microsoft.Extensions.DependencyInjection;
using Wxy.Xaf.RememberLast.Services;

namespace Wxy.Xaf.RememberLast;

/// <summary>
/// XPO 对象属性值记忆模块
/// </summary>
public class WxyXafRememberLastModule : ModuleBase
{
    public WxyXafRememberLastModule()
    {
        // 支持 WinForms 和 Blazor 平台
        RequiredModuleTypes.Add(typeof(DevExpress.ExpressApp.SystemModule.SystemModule));
    }

    public override void Setup(XafApplication application)
    {
        base.Setup(application);

        System.Diagnostics.Debug.WriteLine("[WxyXafRememberLastModule] 模块已加载");

        // 确保服务已注册
        var storageService = application.ServiceProvider?.GetService<ILastValueStorageService>();
        if (storageService == null)
        {
            System.Diagnostics.Debug.WriteLine("[WxyXafRememberLastModule] 警告: ILastValueStorageService 未注册!");
        }
        else
        {
            System.Diagnostics.Debug.WriteLine($"[WxyXafRememberLastModule] ILastValueStorageService 已注册: {storageService.GetType().Name}");
        }
    }

    public override void ExtendModelInterfaces(ModelInterfaceExtenders extenders)
    {
        base.ExtendModelInterfaces(extenders);
        // 如果需要扩展应用程序模型,可以在这里添加
    }
}
