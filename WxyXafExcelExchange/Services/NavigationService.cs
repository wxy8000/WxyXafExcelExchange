#if NET8_0_OR_GREATER
using Microsoft.AspNetCore.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Web;

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// 导航服务实现，使用Blazor NavigationManager进行程序化导航
    /// </summary>
    public class NavigationService : INavigationService
    {
        private readonly NavigationManager _navigationManager;
        private readonly INavigationStateManager _stateManager;

        public NavigationService(NavigationManager navigationManager, INavigationStateManager stateManager = null)
        {
            _navigationManager = navigationManager ?? throw new ArgumentNullException(nameof(navigationManager));
            _stateManager = stateManager;
        }

        /// <summary>
        /// 导航到Excel上传页面
        /// </summary>
        /// <param name="objectTypeName">对象类型名称</param>
        /// <returns>导航任务</returns>
        public async Task NavigateToUploadPageAsync(string objectTypeName)
        {
            if (string.IsNullOrWhiteSpace(objectTypeName))
                throw new ArgumentException("对象类型名称不能为空", nameof(objectTypeName));

            // 保存当前状态（如果有状态管理器）
            if (_stateManager != null)
            {
                await _stateManager.SaveCurrentStateAsync();
            }

            // 构建上传页面URL - 使用路由参数而不是查询参数
            var uploadUrl = $"/excel-upload/{Uri.EscapeDataString(objectTypeName)}";
            
            // 导航到上传页面
            _navigationManager.NavigateTo(uploadUrl);
            
            await Task.CompletedTask;
        }

        /// <summary>
        /// 返回到列表视图页面
        /// </summary>
        /// <param name="objectTypeName">对象类型名称</param>
        /// <returns>导航任务</returns>
        public async Task NavigateBackToListViewAsync(string objectTypeName)
        {
            if (string.IsNullOrWhiteSpace(objectTypeName))
                throw new ArgumentException("对象类型名称不能为空", nameof(objectTypeName));

            // 恢复之前保存的状态（如果有）
            NavigationState savedState = null;
            if (_stateManager != null)
            {
                savedState = await _stateManager.RestoreStateAsync(objectTypeName);
            }

            // 构建列表视图URL
            string listViewUrl;
            if (savedState != null && !string.IsNullOrEmpty(savedState.ReturnUrl))
            {
                // 使用保存的返回URL
                listViewUrl = savedState.ReturnUrl;
            }
            else
            {
                // 默认列表视图URL
                listViewUrl = $"/ListView/{objectTypeName}";
            }

            // 导航回列表视图，强制刷新以显示新导入的数据
            _navigationManager.NavigateTo(listViewUrl, forceLoad: true);
            
            await Task.CompletedTask;
        }

        /// <summary>
        /// 检查是否可以进行导航
        /// </summary>
        /// <returns>如果可以导航返回true，否则返回false</returns>
        public Task<bool> CanNavigateAsync()
        {
            // 检查NavigationManager是否可用
            var canNavigate = _navigationManager != null;
            return Task.FromResult(canNavigate);
        }

        /// <summary>
        /// 导航到指定URL
        /// </summary>
        /// <param name="url">目标URL</param>
        /// <param name="forceLoad">是否强制加载页面</param>
        /// <returns>导航任务</returns>
        public Task NavigateToAsync(string url, bool forceLoad = false)
        {
            if (string.IsNullOrWhiteSpace(url))
                throw new ArgumentException("URL不能为空", nameof(url));

            _navigationManager.NavigateTo(url, forceLoad);
            return Task.CompletedTask;
        }

        /// <summary>
        /// 导航到指定URL并传递参数
        /// </summary>
        /// <param name="url">目标URL</param>
        /// <param name="parameters">URL参数</param>
        /// <param name="forceLoad">是否强制加载页面</param>
        /// <returns>导航任务</returns>
        public Task NavigateToAsync(string url, object parameters, bool forceLoad = false)
        {
            if (string.IsNullOrWhiteSpace(url))
                throw new ArgumentException("URL不能为空", nameof(url));

            var fullUrl = BuildUrlWithParameters(url, parameters);
            _navigationManager.NavigateTo(fullUrl, forceLoad);
            return Task.CompletedTask;
        }

        /// <summary>
        /// 获取当前页面的URL
        /// </summary>
        /// <returns>当前URL</returns>
        public string GetCurrentUrl()
        {
            return _navigationManager.Uri;
        }

        /// <summary>
        /// 获取基础URL
        /// </summary>
        /// <returns>基础URL</returns>
        public string GetBaseUrl()
        {
            return _navigationManager.BaseUri;
        }

        /// <summary>
        /// 构建带参数的URL
        /// </summary>
        /// <param name="baseUrl">基础URL</param>
        /// <param name="parameters">参数对象</param>
        /// <returns>完整的URL</returns>
        public string BuildUrlWithParameters(string baseUrl, object parameters)
        {
            if (string.IsNullOrWhiteSpace(baseUrl))
                throw new ArgumentException("基础URL不能为空", nameof(baseUrl));

            if (parameters == null)
                return baseUrl;

            var queryParams = new List<string>();

            // 使用反射获取参数对象的所有属性
            var properties = parameters.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            
            foreach (var property in properties)
            {
                var value = property.GetValue(parameters);
                if (value != null)
                {
                    var encodedValue = HttpUtility.UrlEncode(value.ToString());
                    queryParams.Add($"{property.Name}={encodedValue}");
                }
            }

            if (queryParams.Any())
            {
                var separator = baseUrl.Contains("?") ? "&" : "?";
                return $"{baseUrl}{separator}{string.Join("&", queryParams)}";
            }

            return baseUrl;
        }
    }

    /// <summary>
    /// 导航状态管理器接口
    /// </summary>
    public interface INavigationStateManager
    {
        /// <summary>
        /// 保存当前导航状态
        /// </summary>
        /// <returns>保存任务</returns>
        Task SaveCurrentStateAsync();

        /// <summary>
        /// 恢复导航状态
        /// </summary>
        /// <param name="objectTypeName">对象类型名称</param>
        /// <returns>导航状态</returns>
        Task<NavigationState> RestoreStateAsync(string objectTypeName);

        /// <summary>
        /// 清除导航状态
        /// </summary>
        /// <param name="objectTypeName">对象类型名称</param>
        /// <returns>清除任务</returns>
        Task ClearStateAsync(string objectTypeName);
    }

    /// <summary>
    /// 导航状态
    /// </summary>
    public class NavigationState
    {
        /// <summary>
        /// 返回URL
        /// </summary>
        public string ReturnUrl { get; set; }

        /// <summary>
        /// 过滤器状态
        /// </summary>
        public Dictionary<string, object> Filters { get; set; } = new Dictionary<string, object>();

        /// <summary>
        /// 排序状态
        /// </summary>
        public Dictionary<string, string> Sorting { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// 分页状态
        /// </summary>
        public PagingState Paging { get; set; } = new PagingState();

        /// <summary>
        /// 自定义状态数据
        /// </summary>
        public Dictionary<string, object> CustomData { get; set; } = new Dictionary<string, object>();
    }

    /// <summary>
    /// 分页状态
    /// </summary>
    public class PagingState
    {
        /// <summary>
        /// 当前页码
        /// </summary>
        public int CurrentPage { get; set; } = 1;

        /// <summary>
        /// 每页记录数
        /// </summary>
        public int PageSize { get; set; } = 20;

        /// <summary>
        /// 总记录数
        /// </summary>
        public int TotalRecords { get; set; }
    }

    /// <summary>
    /// 默认导航状态管理器实现（使用内存存储）
    /// </summary>
    public class InMemoryNavigationStateManager : INavigationStateManager
    {
        private readonly NavigationManager _navigationManager;
        private readonly Dictionary<string, NavigationState> _stateCache = new Dictionary<string, NavigationState>();

        public InMemoryNavigationStateManager(NavigationManager navigationManager)
        {
            _navigationManager = navigationManager ?? throw new ArgumentNullException(nameof(navigationManager));
        }

        /// <summary>
        /// 保存当前导航状态
        /// </summary>
        /// <returns>保存任务</returns>
        public Task SaveCurrentStateAsync()
        {
            var currentUrl = _navigationManager.Uri;
            var uri = new Uri(currentUrl);
            
            // 从URL中提取对象类型
            var queryParams = HttpUtility.ParseQueryString(uri.Query);
            var objectType = queryParams["objectType"];

            if (!string.IsNullOrEmpty(objectType))
            {
                var state = new NavigationState
                {
                    ReturnUrl = currentUrl
                };

                // 保存过滤器和排序信息（如果有）
                foreach (string key in queryParams.AllKeys)
                {
                    if (key != "objectType")
                    {
                        state.CustomData[key] = queryParams[key];
                    }
                }

                _stateCache[objectType] = state;
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// 恢复导航状态
        /// </summary>
        /// <param name="objectTypeName">对象类型名称</param>
        /// <returns>导航状态</returns>
        public Task<NavigationState> RestoreStateAsync(string objectTypeName)
        {
            if (_stateCache.TryGetValue(objectTypeName, out var state))
            {
                return Task.FromResult(state);
            }

            return Task.FromResult<NavigationState>(null);
        }

        /// <summary>
        /// 清除导航状态
        /// </summary>
        /// <param name="objectTypeName">对象类型名称</param>
        /// <returns>清除任务</returns>
        public Task ClearStateAsync(string objectTypeName)
        {
            _stateCache.Remove(objectTypeName);
            return Task.CompletedTask;
        }
    }
}
#endif