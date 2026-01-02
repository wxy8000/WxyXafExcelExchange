using Microsoft.Extensions.DependencyInjection;
using Wxy.Xaf.RememberLast.Services;

namespace Wxy.Xaf.RememberLast;

/// <summary>
/// RememberLast 模块服务注册扩展
/// </summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// 添加 RememberLast 服务
    /// </summary>
    public static IServiceCollection AddRememberLastServices(this IServiceCollection services)
    {
        // 注册内存存储服务
        services.AddSingleton<ILastValueStorageService, MemoryLastValueStorageService>();
        return services;
    }
}
