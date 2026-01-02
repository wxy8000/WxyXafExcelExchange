#if NET8_0_OR_GREATER
using System.Threading.Tasks;

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// 导航服务接口，用于在Blazor应用中进行程序化导航
    /// </summary>
    public interface INavigationService
    {
        /// <summary>
        /// 导航到Excel上传页面
        /// </summary>
        /// <param name="objectTypeName">对象类型名称</param>
        /// <returns>导航任务</returns>
        Task NavigateToUploadPageAsync(string objectTypeName);

        /// <summary>
        /// 返回到列表视图页面
        /// </summary>
        /// <param name="objectTypeName">对象类型名称</param>
        /// <returns>导航任务</returns>
        Task NavigateBackToListViewAsync(string objectTypeName);

        /// <summary>
        /// 检查是否可以进行导航
        /// </summary>
        /// <returns>如果可以导航返回true，否则返回false</returns>
        Task<bool> CanNavigateAsync();

        /// <summary>
        /// 导航到指定URL
        /// </summary>
        /// <param name="url">目标URL</param>
        /// <param name="forceLoad">是否强制加载页面</param>
        /// <returns>导航任务</returns>
        Task NavigateToAsync(string url, bool forceLoad = false);

        /// <summary>
        /// 导航到指定URL并传递参数
        /// </summary>
        /// <param name="url">目标URL</param>
        /// <param name="parameters">URL参数</param>
        /// <param name="forceLoad">是否强制加载页面</param>
        /// <returns>导航任务</returns>
        Task NavigateToAsync(string url, object parameters, bool forceLoad = false);

        /// <summary>
        /// 获取当前页面的URL
        /// </summary>
        /// <returns>当前URL</returns>
        string GetCurrentUrl();

        /// <summary>
        /// 获取基础URL
        /// </summary>
        /// <returns>基础URL</returns>
        string GetBaseUrl();

        /// <summary>
        /// 构建带参数的URL
        /// </summary>
        /// <param name="baseUrl">基础URL</param>
        /// <param name="parameters">参数对象</param>
        /// <returns>完整的URL</returns>
        string BuildUrlWithParameters(string baseUrl, object parameters);
    }
}
#endif