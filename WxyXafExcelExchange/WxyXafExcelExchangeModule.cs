using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.DC;
using DevExpress.Xpo;
using Wxy.Xaf.ExcelExchange.Controllers;
using Wxy.Xaf.ExcelExchange;

namespace Wxy.Xaf.ExcelExchange
{
    [ToolboxItemFilter("Xaf.Platform.Win || Xaf.Platform.Blazor")]
    public sealed class WxyXafExcelExchangeModule : ModuleBase
    {
        public WxyXafExcelExchangeModule()
        {
        }

        public override void Setup(XafApplication application)
        {
            try
            {
                base.Setup(application);

                // 清除配置缓存（确保使用最新的排序逻辑）
                var configManager = new Services.ConfigurationManager();
                configManager.ClearCache();

                // 优先通过入口程序集判断平台类型
                var entryAssembly = Assembly.GetEntryAssembly();
                bool isWinFormsApp = false;
                bool isBlazorApp = false;

                if (entryAssembly != null)
                {
                    var entryAssemblyName = entryAssembly.GetName().Name;

                    // WinForms 应用判断
                    isWinFormsApp = entryAssemblyName?.Contains("Win") == true ||
                                   entryAssemblyName?.Contains("WinForms") == true ||
                                   entryAssemblyName?.EndsWith(".Win") == true;

                    // Blazor 应用判断
                    isBlazorApp = entryAssemblyName?.Contains("Blazor") == true ||
                                  entryAssemblyName?.Contains("Server") == true ||
                                  entryAssemblyName?.EndsWith(".Server") == true;
                }

                // 根据应用类型移除不适合的控制器
                if (isWinFormsApp)
                {
                    // WinForms 应用:移除 Blazor 控制器
                    RemoveControllerType(typeof(BlazorExcelImportExportController));
                }
                else if (isBlazorApp)
                {
                    // Blazor 应用:移除 WinForms 控制器（仅当定义时）
#if NET8_0_WINDOWS || NET462_OR_GREATER
                    RemoveControllerType(typeof(WinFormsExcelImportExportController));
#endif
                    RegisterBlazorJavaScript(application);
                }
                else
                {
                    // 无法判断时,使用程序集检测作为后备方案
                    var platformInfo = DetectPlatform();

                    if (platformInfo.IsWinForms && !platformInfo.IsBlazor)
                    {
                        RemoveControllerType(typeof(BlazorExcelImportExportController));
                    }
                    else if (platformInfo.IsBlazor && !platformInfo.IsWinForms)
                    {
#if NET8_0_WINDOWS || NET462_OR_GREATER
                        RemoveControllerType(typeof(WinFormsExcelImportExportController));
#endif
                        RegisterBlazorJavaScript(application);
                    }
                    // 混合环境且无法判断:保留所有控制器(可能导致重复按钮)
                    // 这种情况下,应用开发者需要手动配置
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WxyXafExcel] 平台检测异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 检测当前平台类型
        /// </summary>
        private (bool IsWinForms, bool IsBlazor) DetectPlatform()
        {
            try
            {
                // 检查WinForms
                var winFormsType = Type.GetType("System.Windows.Forms.Application, System.Windows.Forms");
                var winFormsAssembly = AppDomain.CurrentDomain.GetAssemblies()
                    .FirstOrDefault(a => a.GetName().Name == "System.Windows.Forms");
                bool hasWinForms = winFormsType != null && winFormsAssembly != null;

                // 检查Blazor
                var blazorType = Type.GetType("Microsoft.AspNetCore.Components.ComponentBase, Microsoft.AspNetCore.Components");
                var blazorAssembly = AppDomain.CurrentDomain.GetAssemblies()
                    .FirstOrDefault(a => a.GetName().Name?.Contains("Blazor") == true ||
                                        a.GetName().Name?.Contains("AspNetCore") == true);
                bool hasBlazor = blazorType != null && blazorAssembly != null;

                return (hasWinForms, hasBlazor);
            }
            catch (Exception)
            {
                return (false, false);
            }
        }

        /// <summary>
        /// 注册Blazor JavaScript文件
        /// </summary>
        private void RegisterBlazorJavaScript(XafApplication application)
        {
            try
            {
                // 注意：在ASP.NET Core Blazor中，静态资源文件通过wwwroot文件夹自动提供
                // 项目文件已配置 StaticWebBasePath 和 StaticWebAsset，无需手动注册
                // JavaScript和CSS文件会自动通过 _content/WxyXafExcel/ 路径访问

                // _content/WxyXafExcel/js/download.js - JavaScript文件
                // _content/WxyXafExcel/css/disable-dictionary-links.css - CSS文件

                System.Diagnostics.Debug.WriteLine("[WxyXafExcel] Blazor静态资源已通过wwwroot自动配置");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WxyXafExcel] 注册Blazor资源时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 移除指定类型的控制器
        /// </summary>
        private void RemoveControllerType(Type controllerType)
        {
            try
            {
                // 从控制器类型列表中移除
                var controllerTypesField = typeof(ModuleBase).GetField("controllerTypes",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

                if (controllerTypesField != null)
                {
                    var controllerTypes = controllerTypesField.GetValue(this) as IList<Type>;
                    if (controllerTypes != null && controllerTypes.Contains(controllerType))
                    {
                        controllerTypes.Remove(controllerType);
                    }
                }
            }
            catch (Exception)
            {
                // 忽略控制器移除失败
            }
        }

        public override void CustomizeTypesInfo(ITypesInfo typesInfo)
        {
            try
            {
                base.CustomizeTypesInfo(typesInfo);
            }
            catch (Exception ex)
            {
                // 忽略XPO重复注册错误，这可能是由于用户项目中使用了相同命名空间导致的
                if (!ex.Message.Contains("Dictionary already contains ClassInfo") && !ex.Message.Contains("already loaded twice"))
                {
                    // 其他错误也忽略，避免模块加载失败
                }
            }
        }
    }
}