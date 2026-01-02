using System;
using Microsoft.Extensions.DependencyInjection;
using Wxy.Xaf.ExcelExchange.Services;

#if NET8_0_OR_GREATER
using Microsoft.JSInterop;
using Microsoft.AspNetCore.Components;
#endif

namespace Wxy.Xaf.ExcelExchange.Extensions
{
    /// <summary>
    /// 服务集合扩展方法
    /// </summary>
    public static class ServiceCollectionExtensions
    {
        /// <summary>
        /// 添加Excel导入导出模块服务
        /// </summary>
        /// <param name="services">服务集合</param>
        /// <returns>服务集合</returns>
        public static IServiceCollection AddXafExcelImportExport(this IServiceCollection services)
        {
            // 注册核心Excel服务
            services.AddScoped<IExcelImportExportService, ExcelImportExportService>();
            
            // 注册导入服务（向后兼容）
            services.AddScoped<IImportService, ImportService>();
            
            // 注册对话框导入服务（向后兼容）
            services.AddScoped<IDialogExcelImportService, DefaultExcelImportService>();
            
            // 注意：导航服务已移动到Blazor特定的注册方法中，因为它们依赖于NavigationManager

            return services;
        }

        /// <summary>
        /// 添加WinForms平台的Excel导入导出服务
        /// </summary>
        /// <param name="services">服务集合</param>
        /// <returns>服务集合</returns>
        public static IServiceCollection AddXafExcelImportExportForWinForms(this IServiceCollection services)
        {
            // 注册基础服务
            services.AddXafExcelImportExport();

            // 注册WinForms平台特定服务
#if NET8_0_WINDOWS || NET462_OR_GREATER
            services.AddScoped<IPlatformFileService>(provider =>
            {
                var application = provider.GetService<DevExpress.ExpressApp.XafApplication>();
                return new WinFormsPlatformFileService(application);
            });
#else
            // 非Windows平台不支持WinForms服务
            services.AddScoped<IPlatformFileService>(provider =>
            {
                throw new PlatformNotSupportedException("WinForms platform file service is only available on Windows. Please use Blazor platform instead.");
            });
#endif

            return services;
        }

        /// <summary>
        /// 添加Blazor平台的Excel导入导出服务
        /// </summary>
        /// <param name="services">服务集合</param>
        /// <returns>服务集合</returns>
        public static IServiceCollection AddXafExcelImportExportForBlazor(this IServiceCollection services)
        {
            // 注册基础服务
            services.AddXafExcelImportExport();
            
#if NET8_0_OR_GREATER
            // 注册Blazor特定的导航服务（依赖于NavigationManager）
            services.AddScoped<INavigationService, NavigationService>();
            services.AddScoped<INavigationStateManager, InMemoryNavigationStateManager>();
            
            // 注册Blazor平台特定服务
            services.AddScoped<IPlatformFileService>(provider =>
            {
                var application = provider.GetService<DevExpress.ExpressApp.XafApplication>();
                
                try
                {
                    var jsRuntime = provider.GetService<IJSRuntime>();
                    var navigationManager = provider.GetService<NavigationManager>();
                    
                    if (jsRuntime != null && navigationManager != null)
                    {
                        return new BlazorPlatformFileService(application, jsRuntime, navigationManager);
                    }
                }
                catch
                {
                    // 如果获取Blazor服务失败，使用基础构造函数
                }
                
                return new BlazorPlatformFileService(application);
            });
#else
            // .NET Framework版本的降级实现
            services.AddScoped<IPlatformFileService>(provider =>
            {
                var application = provider.GetService<DevExpress.ExpressApp.XafApplication>();
                return new BlazorPlatformFileService(application);
            });
#endif

            return services;
        }

        /// <summary>
        /// 添加Excel导入导出模块服务（带配置选项）
        /// </summary>
        /// <param name="services">服务集合</param>
        /// <param name="configureOptions">配置选项委托</param>
        /// <returns>服务集合</returns>
        public static IServiceCollection AddXafExcelImportExport(this IServiceCollection services, 
            System.Action<ExcelImportExportOptions> configureOptions)
        {
#if NET8_0_OR_GREATER
            // 配置选项（仅限.NET 8.0+）
            services.Configure(configureOptions);
#endif
            
            // 注册基础服务
            services.AddXafExcelImportExport();

            return services;
        }

        /// <summary>
        /// 添加WinForms平台的Excel导入导出服务（带配置选项）
        /// </summary>
        /// <param name="services">服务集合</param>
        /// <param name="configureOptions">配置选项委托</param>
        /// <returns>服务集合</returns>
        public static IServiceCollection AddXafExcelImportExportForWinForms(this IServiceCollection services, 
            System.Action<ExcelImportExportOptions> configureOptions)
        {
            services.AddXafExcelImportExport(configureOptions);
            services.AddXafExcelImportExportForWinForms();
            return services;
        }

        /// <summary>
        /// 添加Blazor平台的Excel导入导出服务（带配置选项）
        /// </summary>
        /// <param name="services">服务集合</param>
        /// <param name="configureOptions">配置选项委托</param>
        /// <returns>服务集合</returns>
        public static IServiceCollection AddXafExcelImportExportForBlazor(this IServiceCollection services, 
            System.Action<ExcelImportExportOptions> configureOptions)
        {
            services.AddXafExcelImportExport(configureOptions);
            services.AddXafExcelImportExportForBlazor();
            return services;
        }
    }

    /// <summary>
    /// Excel导入导出模块配置选项
    /// </summary>
    public class ExcelImportExportOptions
    {
        /// <summary>
        /// 是否创建缺失的引用对象
        /// </summary>
        public bool CreateMissingReferences { get; set; } = false;

        /// <summary>
        /// 最大文件大小（字节）
        /// </summary>
        public long MaxFileSize { get; set; } = 10 * 1024 * 1024; // 10MB

        /// <summary>
        /// 支持的文件格式
        /// </summary>
        public string[] SupportedFormats { get; set; } = { ".csv", ".xlsx", ".xls" };

        /// <summary>
        /// 默认导入选项
        /// </summary>
        public ImportOptions DefaultImportOptions { get; set; } = new ImportOptions();

        /// <summary>
        /// 上传页面路径
        /// </summary>
        public string UploadPagePath { get; set; } = "/ExcelUpload";

        /// <summary>
        /// 列表视图页面路径模板
        /// </summary>
        public string ListViewPathTemplate { get; set; } = "/ListView/{0}";
    }
}