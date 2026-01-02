using System;
using System.Linq;
using System.Threading.Tasks;

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// 平台特定文件服务接口
    /// </summary>
    public interface IPlatformFileService
    {
        /// <summary>
        /// 显示文件选择对话框
        /// </summary>
        /// <param name="options">文件选择选项</param>
        /// <returns>文件选择结果</returns>
        Task<FileSelectionResult> ShowFileSelectionDialogAsync(FileSelectionOptions options);

        /// <summary>
        /// 选择文件（简化的便捷方法）
        /// </summary>
        /// <param name="allowedExtensions">允许的文件扩展名数组（如 ".xlsx", ".csv"）</param>
        /// <returns>文件选择结果，包含单个文件的文件名和内容</returns>
        Task<FileSelectionResult> PickFileAsync(string[] allowedExtensions = null);

        /// <summary>
        /// 显示文件保存对话框
        /// </summary>
        /// <param name="options">文件保存选项</param>
        /// <returns>文件保存结果</returns>
        Task<FileSaveResult> ShowFileSaveDialogAsync(FileSaveOptions options);

        /// <summary>
        /// 下载文件到客户端
        /// </summary>
        /// <param name="fileContent">文件内容</param>
        /// <param name="fileName">文件名</param>
        /// <param name="mimeType">MIME类型</param>
        /// <returns>下载结果</returns>
        Task<FileDownloadResult> DownloadFileAsync(byte[] fileContent, string fileName, string mimeType);

        /// <summary>
        /// 显示消息
        /// </summary>
        /// <param name="message">消息内容</param>
        /// <param name="title">标题</param>
        /// <param name="messageType">消息类型</param>
        /// <returns>任务</returns>
        Task ShowMessageAsync(string message, string title = null, MessageType messageType = MessageType.Information);

        /// <summary>
        /// 显示确认对话框
        /// </summary>
        /// <param name="message">消息内容</param>
        /// <param name="title">标题</param>
        /// <returns>用户选择结果</returns>
        Task<bool> ShowConfirmationAsync(string message, string title = null);

        /// <summary>
        /// 显示导入选项对话框
        /// </summary>
        /// <param name="title">对话框标题</param>
        /// <param name="defaultMode">默认导入模式</param>
        /// <returns>用户选择的导入模式,如果取消则返回null</returns>
        Task<ExcelImportMode?> ShowImportOptionsDialogAsync(string title = null, ExcelImportMode? defaultMode = null);

        /// <summary>
        /// 导航到指定页面
        /// </summary>
        /// <param name="url">目标URL</param>
        /// <param name="newWindow">是否在新窗口打开</param>
        /// <returns>导航结果</returns>
        Task<NavigationResult> NavigateToAsync(string url, bool newWindow = false);
    }

    /// <summary>
    /// 文件选择选项
    /// </summary>
    public class FileSelectionOptions
    {
        /// <summary>
        /// 文件过滤器
        /// </summary>
        public string Filter { get; set; } = "Excel文件 (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|所有文件 (*.*)|*.*";

        /// <summary>
        /// 对话框标题
        /// </summary>
        public string Title { get; set; } = "选择要导入的文件";

        /// <summary>
        /// 是否允许多选
        /// </summary>
        public bool Multiselect { get; set; } = false;

        /// <summary>
        /// 初始目录
        /// </summary>
        public string InitialDirectory { get; set; }

        /// <summary>
        /// 默认文件名
        /// </summary>
        public string DefaultFileName { get; set; }
    }

    /// <summary>
    /// 文件保存选项
    /// </summary>
    public class FileSaveOptions
    {
        /// <summary>
        /// 文件过滤器（xlsx排在第一位，作为默认格式）
        /// </summary>
        public string Filter { get; set; } = "Excel文件 (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls|CSV文件 (*.csv)|*.csv|所有文件 (*.*)|*.*";

        /// <summary>
        /// 对话框标题
        /// </summary>
        public string Title { get; set; } = "保存文件";

        /// <summary>
        /// 默认文件名
        /// </summary>
        public string DefaultFileName { get; set; }

        /// <summary>
        /// 默认扩展名（xlsx）
        /// </summary>
        public string DefaultExtension { get; set; } = "xlsx";

        /// <summary>
        /// 是否覆盖提示
        /// </summary>
        public bool OverwritePrompt { get; set; } = true;

        /// <summary>
        /// 是否自动添加扩展名
        /// </summary>
        public bool AddExtension { get; set; } = true;

        /// <summary>
        /// 初始目录
        /// </summary>
        public string InitialDirectory { get; set; }
    }

    /// <summary>
    /// 文件选择结果
    /// </summary>
    public class FileSelectionResult
    {
        /// <summary>
        /// 是否成功选择
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 选择的文件路径
        /// </summary>
        public string[] FilePaths { get; set; } = new string[0];

        /// <summary>
        /// 文件内容（适用于Web平台）
        /// </summary>
        public FileContentInfo[] FileContents { get; set; } = new FileContentInfo[0];

        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 是否用户取消
        /// </summary>
        public bool IsCancelled { get; set; }

        /// <summary>
        /// 获取第一个文件的文件名（便捷属性）
        /// </summary>
        public string FileName => FileContents?.FirstOrDefault()?.FileName;

        /// <summary>
        /// 获取第一个文件的内容（便捷属性）
        /// </summary>
        public byte[] FileContent => FileContents?.FirstOrDefault()?.Content;
    }

    /// <summary>
    /// 文件保存结果
    /// </summary>
    public class FileSaveResult
    {
        /// <summary>
        /// 是否成功保存
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 保存的文件路径
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 是否用户取消
        /// </summary>
        public bool IsCancelled { get; set; }
    }

    /// <summary>
    /// 文件下载结果
    /// </summary>
    public class FileDownloadResult
    {
        /// <summary>
        /// 是否成功下载
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 导航结果
    /// </summary>
    public class NavigationResult
    {
        /// <summary>
        /// 是否成功导航
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// 文件内容信息
    /// </summary>
    public class FileContentInfo
    {
        /// <summary>
        /// 文件名
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 文件内容
        /// </summary>
        public byte[] Content { get; set; }

        /// <summary>
        /// 文件大小
        /// </summary>
        public long Size { get; set; }

        /// <summary>
        /// MIME类型
        /// </summary>
        public string MimeType { get; set; }
    }

    /// <summary>
    /// 消息类型
    /// </summary>
    public enum MessageType
    {
        /// <summary>
        /// 信息
        /// </summary>
        Information,

        /// <summary>
        /// 警告
        /// </summary>
        Warning,

        /// <summary>
        /// 错误
        /// </summary>
        Error,

        /// <summary>
        /// 成功
        /// </summary>
        Success
    }
}