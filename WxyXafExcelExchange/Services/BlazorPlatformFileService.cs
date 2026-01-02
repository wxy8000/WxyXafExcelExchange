using System;
using System.Linq;
using System.Threading.Tasks;
using DevExpress.ExpressApp;

#if NET8_0_OR_GREATER
using Microsoft.JSInterop;
using Microsoft.AspNetCore.Components;
#endif

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// Blazor平台文件服务实现
    /// </summary>
    public class BlazorPlatformFileService : IPlatformFileService
    {
        private readonly XafApplication _application;
#if NET8_0_OR_GREATER
        private readonly IJSRuntime _jsRuntime;
        private readonly NavigationManager _navigationManager;

        public BlazorPlatformFileService(XafApplication application, IJSRuntime jsRuntime, NavigationManager navigationManager)
        {
            _application = application; // 允许为null，在使用时再检查
            _jsRuntime = jsRuntime ?? throw new ArgumentNullException(nameof(jsRuntime));
            _navigationManager = navigationManager ?? throw new ArgumentNullException(nameof(navigationManager));
        }

        public BlazorPlatformFileService(XafApplication application)
        {
            _application = application; // 允许为null，在使用时再检查
            _jsRuntime = null;
            _navigationManager = null;
        }
#else
        public BlazorPlatformFileService(XafApplication application)
        {
            _application = application; // 允许为null，在使用时再检查
        }
#endif

        /// <summary>
        /// 显示文件选择对话框
        /// </summary>
        public async Task<FileSelectionResult> ShowFileSelectionDialogAsync(FileSelectionOptions options)
        {
            var result = new FileSelectionResult();

#if NET8_0_OR_GREATER
            try
            {
                // 在Blazor中，文件选择通常通过导航到专门的文件上传页面实现
                // 这里返回一个指示需要导航的结果
                result.IsSuccess = false;
                result.ErrorMessage = "请使用文件上传页面选择文件";
                return result;
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.ErrorMessage = $"文件选择失败: {ex.Message}";
                return result;
            }
#else
            result.IsSuccess = false;
            result.ErrorMessage = "当前.NET Framework版本不支持Blazor文件选择";
            return result;
#endif
        }

        /// <summary>
        /// 选择文件（简化的便捷方法）
        /// </summary>
        public async Task<FileSelectionResult> PickFileAsync(string[] allowedExtensions = null)
        {
            var result = new FileSelectionResult { IsCancelled = true };

#if NET8_0_OR_GREATER
            try
            {
                if (_jsRuntime == null)
                {
                    result.ErrorMessage = "JavaScript运行时不可用";
                    return result;
                }

                // 构建accept属性值
                string accept = "*/*";
                if (allowedExtensions != null && allowedExtensions.Length > 0)
                {
                    accept = string.Join(",", allowedExtensions);
                }

                // 使用JavaScript创建文件输入并触发选择
                var script = $@"
                    (function() {{
                        return new Promise((resolve, reject) => {{
                            try {{
                                // 创建文件输入元素
                                const input = document.createElement('input');
                                input.type = 'file';
                                input.accept = '{EscapeJavaScriptString(accept)}';
                                input.style.display = 'none';

                                // 处理文件选择
                                input.onchange = async (e) => {{
                                    const file = e.target.files[0];
                                    if (!file) {{
                                        document.body.removeChild(input);
                                        resolve({{ cancelled: true }});
                                        return;
                                    }}

                                    try {{
                                        // 读取文件内容
                                        const reader = new FileReader();
                                        reader.onload = (event) => {{
                                            const arrayBuffer = event.target.result;
                                            const uint8Array = new Uint8Array(arrayBuffer);
                                            const base64 = btoa(String.fromCharCode.apply(null, uint8Array));

                                            document.body.removeChild(input);
                                            resolve({{
                                                cancelled: false,
                                                fileName: file.name,
                                                fileSize: file.size,
                                                mimeType: file.type,
                                                base64: base64
                                            }});
                                        }};
                                        reader.onerror = () => {{
                                            document.body.removeChild(input);
                                            resolve({{ cancelled: true, error: '文件读取失败' }});
                                        }};
                                        reader.readAsArrayBuffer(file);
                                    }} catch (err) {{
                                        document.body.removeChild(input);
                                        resolve({{ cancelled: true, error: err.toString() }});
                                    }}
                                }};

                                // 处理取消
                                input.oncancel = () => {{
                                    document.body.removeChild(input);
                                    resolve({{ cancelled: true }});
                                }};

                                // 添加到DOM并触发点击
                                document.body.appendChild(input);
                                input.click();
                            }} catch (err) {{
                                reject(err.toString());
                            }}
                        }});
                    }})();
                ";

                var fileInfo = await _jsRuntime.InvokeAsync<FileInfoDto>("eval", script);

                if (fileInfo == null || fileInfo.Cancelled)
                {
                    result.IsCancelled = true;
                    return result;
                }

                if (!string.IsNullOrEmpty(fileInfo.Error))
                {
                    result.ErrorMessage = fileInfo.Error;
                    return result;
                }

                // 转换Base64为字节数组
                var fileBytes = Convert.FromBase64String(fileInfo.Base64);

                result.IsSuccess = true;
                result.IsCancelled = false;
                result.FileContents = new[]
                {
                    new FileContentInfo
                    {
                        FileName = fileInfo.FileName,
                        Content = fileBytes,
                        Size = fileInfo.FileSize,
                        MimeType = fileInfo.MimeType
                    }
                };
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.IsCancelled = true;
                result.ErrorMessage = $"文件选择失败: {ex.Message}";
            }
#else
            result.ErrorMessage = "当前版本不支持文件选择";
#endif

            return result;
        }

        /// <summary>
        /// 文件信息传输对象（用于JS互操作）
        /// </summary>
        private class FileInfoDto
        {
            public bool Cancelled { get; set; }
            public string FileName { get; set; }
            public long FileSize { get; set; }
            public string MimeType { get; set; }
            public string Base64 { get; set; }
            public string Error { get; set; }
        }

        /// <summary>
        /// 显示文件保存对话框（Blazor中不适用，直接下载）
        /// </summary>
        public async Task<FileSaveResult> ShowFileSaveDialogAsync(FileSaveOptions options)
        {
            var result = new FileSaveResult();

#if NET8_0_OR_GREATER
            try
            {
                // 在Blazor中，文件保存通过下载实现
                result.IsSuccess = false;
                result.ErrorMessage = "请使用下载功能保存文件";
                return result;
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.ErrorMessage = $"文件保存对话框失败: {ex.Message}";
                return result;
            }
#else
            result.IsSuccess = false;
            result.ErrorMessage = "当前.NET Framework版本不支持Blazor文件保存";
            return result;
#endif
        }

        /// <summary>
        /// 下载文件到客户端
        /// </summary>
        public async Task<FileDownloadResult> DownloadFileAsync(byte[] fileContent, string fileName, string mimeType)
        {
            var result = new FileDownloadResult();

#if NET8_0_OR_GREATER
            try
            {
                if (_jsRuntime == null)
                {
                    result.ErrorMessage = "JavaScript运行时不可用";
                    return result;
                }

                // 转换为Base64
                var base64Content = Convert.ToBase64String(fileContent);
                
                // 安全的文件名转义
                var safeFileName = EscapeJavaScriptString(fileName);
                
                // 创建下载脚本
                var downloadScript = $@"
                    (function() {{
                        try {{
                            console.log('Blazor文件下载开始: {safeFileName}');
                            
                            // 转换Base64为Blob
                            const binaryString = window.atob('{base64Content}');
                            const bytes = new Uint8Array(binaryString.length);
                            for (let i = 0; i < binaryString.length; i++) {{
                                bytes[i] = binaryString.charCodeAt(i);
                            }}
                            const blob = new Blob([bytes], {{ type: '{mimeType}' }});
                            
                            // 创建下载链接
                            const url = URL.createObjectURL(blob);
                            const a = document.createElement('a');
                            a.href = url;
                            a.download = '{safeFileName}';
                            a.style.display = 'none';
                            
                            // 添加到DOM并触发下载
                            document.body.appendChild(a);
                            a.click();
                            
                            // 清理资源
                            setTimeout(() => {{
                                document.body.removeChild(a);
                                URL.revokeObjectURL(url);
                                console.log('Blazor文件下载完成: {safeFileName}');
                            }}, 100);
                            
                            return true;
                        }} catch (error) {{
                            console.error('Blazor文件下载失败:', error);
                            return false;
                        }}
                    }})();
                ";

                // 执行JavaScript
                var downloadSuccess = await _jsRuntime.InvokeAsync<bool>("eval", downloadScript);
                
                if (downloadSuccess)
                {
                    result.IsSuccess = true;
                }
                else
                {
                    result.ErrorMessage = "JavaScript下载执行失败";
                }
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.ErrorMessage = $"文件下载失败: {ex.Message}";
            }
#else
            result.IsSuccess = false;
            result.ErrorMessage = "当前.NET Framework版本不支持Blazor文件下载";
#endif

            return result;
        }

        /// <summary>
        /// 显示消息
        /// </summary>
        public async Task ShowMessageAsync(string message, string title = null, MessageType messageType = MessageType.Information)
        {
#if NET8_0_OR_GREATER
            try
            {
                if (_jsRuntime == null)
                {
                    // 降级到XAF消息显示
                    _application?.ShowViewStrategy?.ShowMessage(message);
                    return;
                }

                var icon = messageType switch
                {
                    MessageType.Warning => "⚠️",
                    MessageType.Error => "❌",
                    MessageType.Success => "✅",
                    _ => "ℹ️"
                };

                var titleIcon = messageType switch
                {
                    MessageType.Warning => "⚠️ 警告",
                    MessageType.Error => "❌ 错误",
                    MessageType.Success => "✅ 成功",
                    _ => "ℹ️ 提示"
                };

                var fullTitle = string.IsNullOrEmpty(title) ? titleIcon : titleIcon + " - " + title;
                var fullMessage = message.Replace("\n", "<br>").Replace(" ", "&nbsp;");

                // 转义JavaScript字符串
                var escapedTitle = EscapeJavaScriptString(fullTitle);
                var escapedMessageBody = EscapeJavaScriptString(fullMessage);
                var escapedTitleForCopy = EscapeJavaScriptString(title ?? "");
                var escapedMessageForCopy = EscapeJavaScriptString(message);

                // 检查是否为导入完成消息,使用固定宽度和特殊样式
                var isImportMessage = (messageType == MessageType.Success || messageType == MessageType.Warning) &&
                                     !string.IsNullOrEmpty(title) &&
                                     (title.Contains("导入") || title.Contains("Import"));

                var dialogWidth = isImportMessage ? "width: 680px;" : "max-width: 600px;";
                var dialogBorderRadius = isImportMessage ? "border-radius: 16px;" : "border-radius: 8px;";
                var dialogPadding = isImportMessage ? "padding: 0;" : "padding: 24px;";
                var dialogShadow = isImportMessage ? "box-shadow: 0 20px 60px rgba(0,0,0,0.3);" : "box-shadow: 0 4px 20px rgba(0,0,0,0.3);";
                var headerBg = isImportMessage ? "background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);" : "";

                // 创建可选择的HTML对话框(使用字符串拼接避免箭头函数语法问题)
                var script = "(function() { try { " +
                    "const existingDialog = document.getElementById('xaf-selectable-dialog'); " +
                    "if (existingDialog) { document.body.removeChild(existingDialog); } " +
                    "const dialog = document.createElement('div'); " +
                    "dialog.id = 'xaf-selectable-dialog'; " +
                    "dialog.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.6); display: flex; align-items: center; justify-content: center; z-index: 99999; font-family: system-ui, -apple-system, sans-serif;'; " +
                    "const content = document.createElement('div'); " +
                    "content.style.cssText = 'background: white; " + dialogBorderRadius + " " + dialogPadding + " " + dialogWidth + " max-height: 75vh; overflow: hidden; display: flex; flex-direction: column; " + dialogShadow + " animation: dialogSlideIn 0.3s ease-out;'; " +

                    // 添加动画样式
                    "if (!document.getElementById('xaf-dialog-animations')) { " +
                        "const style = document.createElement('style'); " +
                        "style.id = 'xaf-dialog-animations'; " +
                        "style.textContent = '@keyframes dialogSlideIn { from { opacity: 0; transform: scale(0.95) translateY(-20px); } to { opacity: 1; transform: scale(1) translateY(0); } }'; " +
                        "document.head.appendChild(style); " +
                    "} " +

                    // 如果是导入消息,创建头部
                    (isImportMessage ?
                        ("const header = document.createElement('div'); " +
                        "header.style.cssText = '" + headerBg + " color: white; padding: 20px 24px; display: flex; align-items: center; gap: 12px; flex-shrink: 0;'; " +
                        "header.innerHTML = '<span style=\\'font-size: 32px;\\'>" + icon + "</span><div><div style=\\'font-size: 20px; font-weight: 600; margin-bottom: 4px;\\'>" + escapedTitle + "</div><div style=\\'font-size: 13px; opacity: 0.9;\\'>导入操作已完成</div></div>'; " +
                        "content.appendChild(header); ")
                        : "const header = null; ") +

                    "const titleElement = document.createElement('div'); " +
                    "titleElement.style.cssText = '" + (isImportMessage ? "display: none;" : "font-size: 18px; font-weight: bold; margin-bottom: 16px; color: #333; display: flex; align-items: center; gap: 8px;") + "'; " +
                    "titleElement.innerHTML = '" + icon + " " + escapedTitle + "'; " +
                    "const messageElement = document.createElement('div'); " +
                    "messageElement.style.cssText = '" + (isImportMessage ? "flex: 1; overflow-y: auto; padding: 20px 24px;" : "color: #555; line-height: 1.6; white-space: pre-wrap;") + " user-select: text; -webkit-user-select: text; -moz-user-select: text; -ms-user-select: text; font-size: 14px;" + (isImportMessage ? "" : " margin-bottom: 20px;") + "'; " +
                    "messageElement.innerHTML = '" + escapedMessageBody + "'; " +
                    "const buttonContainer = document.createElement('div'); " +
                    "buttonContainer.style.cssText = '" + (isImportMessage ? "padding: 16px 24px; border-top: 1px solid #e8e8e8; display: flex; gap: 12px; justify-content: flex-end; background: #fafafa; flex-shrink: 0;" : "display: flex; gap: 12px; justify-content: flex-end; margin-bottom: 20px;") + "'; " +
                    "const copyButton = document.createElement('button'); " +
                    "copyButton.textContent = '📋 复制'; " +
                    "copyButton.style.cssText = '" + (isImportMessage ? "padding: 11px 20px;" : "padding: 10px 24px;") + " background: #6c757d; color: white; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.2s;'; " +
                    "copyButton.onmouseover = function() { copyButton.style.background = '#5a6268'; copyButton.style.transform = 'translateY(-1px)'; }; " +
                    "copyButton.onmouseout = function() { copyButton.style.background = '#6c757d'; copyButton.style.transform = 'translateY(0)'; }; " +
                    "copyButton.onclick = function() { " +
                        "const textToCopy = '" + escapedTitleForCopy + "\\n\\n" + escapedMessageForCopy + "'; " +
                        "navigator.clipboard.writeText(textToCopy).then(function() { " +
                            "copyButton.textContent = '✅ 已复制'; " +
                            "setTimeout(function() { copyButton.textContent = '📋 复制'; }, 2000); " +
                        "}).catch(function() { " +
                            "copyButton.textContent = '❌ 复制失败'; " +
                            "setTimeout(function() { copyButton.textContent = '📋 复制'; }, 2000); " +
                        "}); " +
                    "}; " +
                    "const okButton = document.createElement('button'); " +
                    "okButton.textContent = '确定'; " +
                    "okButton.style.cssText = '" + (isImportMessage ? "padding: 11px 32px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);" : "padding: 10px 24px; background: #0078d4;") + " color: white; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 600; transition: all 0.2s;'; " +
                    "okButton.onmouseover = function() { okButton.style.transform = '" + (isImportMessage ? "translateY(-2px); box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);" : "translateY(-1px);") + "' }; " +
                    (isImportMessage ? "" : "okButton.onmouseover = function() { okButton.style.background = '#106ebe'; }; ") +
                    "okButton.onmouseout = function() { okButton.style.transform = 'translateY(0)'; " + (isImportMessage ? "okButton.style.boxShadow = 'none';" : "okButton.style.background = '#0078d4';") + " }; " +
                    "okButton.onclick = function() { document.body.removeChild(dialog); }; " +
                    "buttonContainer.appendChild(copyButton); " +
                    "buttonContainer.appendChild(okButton); " +
                    "if (header) { content.appendChild(header); } " +
                    "content.appendChild(titleElement); " +
                    "content.appendChild(buttonContainer); " +
                    "content.appendChild(messageElement); " +
                    "dialog.appendChild(content); " +
                    "document.body.appendChild(dialog); " +
                    "dialog.onclick = function(e) { if (e.target === dialog) { document.body.removeChild(dialog); } }; " +
                    "const escapeHandler = function(e) { if (e.key === 'Escape') { document.removeEventListener('keydown', escapeHandler); const currentDialog = document.getElementById('xaf-selectable-dialog'); if (currentDialog) { document.body.removeChild(currentDialog); } } }; " +
                    "document.addEventListener('keydown', escapeHandler); " +
                    "return true; " +
                    "} catch (error) { console.error('显示可选文本对话框失败:', error); return false; } " +
                    "})();";

                var success = await _jsRuntime.InvokeAsync<bool>("eval", script);
                if (!success)
                {
                    // 降级到标准alert
                    await _jsRuntime.InvokeVoidAsync("alert", icon + " " + title + "\n\n" + message);
                }
            }
            catch (Exception ex)
            {
                // 降级到XAF消息显示
                _application?.ShowViewStrategy?.ShowMessage(message + "\n\n(消息显示异常: " + ex.Message + ")");
            }
#else
            // 降级到XAF消息显示
            _application?.ShowViewStrategy?.ShowMessage(message);
#endif
        }

        /// <summary>
        /// 显示确认对话框
        /// </summary>
        public async Task<bool> ShowConfirmationAsync(string message, string title = null)
        {
#if NET8_0_OR_GREATER
            try
            {
                if (_jsRuntime == null)
                {
                    // 降级到XAF消息显示
                    _application?.ShowViewStrategy?.ShowMessage(message);
                    return false;
                }

                var fullMessage = string.IsNullOrEmpty(title) ? message : $"{title}\n\n{message}";
                // 直接传递消息,不转义,因为 IJSRuntime 会自动处理字符串转义
                return await _jsRuntime.InvokeAsync<bool>("confirm", fullMessage);
            }
            catch (Exception ex)
            {
                // 降级到XAF消息显示
                _application?.ShowViewStrategy?.ShowMessage($"{message}\n\n(确认对话框异常: {ex.Message})");
                return false;
            }
#else
            // 降级到XAF消息显示
            _application?.ShowViewStrategy?.ShowMessage(message);
            return false;
#endif
        }

        /// <summary>
        /// 显示导入选项对话框
        /// </summary>
        public async Task<ExcelImportMode?> ShowImportOptionsDialogAsync(string title = null, ExcelImportMode? defaultMode = null)
        {
#if NET8_0_OR_GREATER
            try
            {
                if (_jsRuntime == null)
                {
                    // 降级: 返回默认模式
                    return defaultMode ?? ExcelImportMode.CreateOrUpdate;
                }

                var dialogTitle = string.IsNullOrEmpty(title) ? "选择导入模式" : title;
                var escapedTitle = EscapeJavaScriptString(dialogTitle);

                // 确定默认选中的模式
                var defaultIndex = defaultMode.HasValue ? (int)defaultMode.Value : 2; // 默认 CreateOrUpdate = 2

                // 创建对话框脚本
                var script = $@"
                    (function() {{
                        try {{
                            // 移除已存在的对话框
                            const existingDialog = document.getElementById('xaf-import-options-dialog');
                            if (existingDialog) {{
                                document.body.removeChild(existingDialog);
                            }}

                            // 创建对话框容器
                            const dialog = document.createElement('div');
                            dialog.id = 'xaf-import-options-dialog';
                            dialog.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); display: flex; align-items: center; justify-content: center; z-index: 99999; font-family: system-ui, -apple-system, sans-serif;';

                            // 创建内容容器
                            const content = document.createElement('div');
                            content.style.cssText = 'background: white; border-radius: 8px; padding: 24px; max-width: 500px; box-shadow: 0 4px 20px rgba(0,0,0,0.3);';

                            // 标题
                            const title = document.createElement('div');
                            title.style.cssText = 'font-size: 18px; font-weight: bold; margin-bottom: 16px; color: #333;';
                            title.textContent = '{escapedTitle}';
                            content.appendChild(title);

                            // 说明文本
                            const description = document.createElement('div');
                            description.style.cssText = 'color: #666; margin-bottom: 20px; line-height: 1.5;';
                            description.textContent = '请选择数据导入模式:';
                            content.appendChild(description);

                            // 选项列表
                            const optionsContainer = document.createElement('div');
                            optionsContainer.style.cssText = 'display: flex; flex-direction: column; gap: 12px; margin-bottom: 24px;';

                            // 创建选项
                            const options = [
                                {{ value: 0, label: '仅新增 (Insert)', desc: '只创建新记录,如果记录已存在则跳过' }},
                                {{ value: 1, label: '仅更新 (Update)', desc: '只更新现有记录,如果记录不存在则跳过' }},
                                {{ value: 2, label: '新增或更新 (InsertOrUpdate)', desc: '存在则更新,不存在则新增 (推荐)' }},
                                {{ value: 3, label: '替换全部 (ReplaceAll)', desc: '删除所有现有记录后重新导入' }}
                            ];

                            options.forEach((option, index) => {{
                                const optionDiv = document.createElement('div');
                                optionDiv.style.cssText = 'display: flex; align-items: flex-start; gap: 10px; padding: 10px; border: 2px solid #e0e0e0; border-radius: 6px; cursor: pointer; transition: all 0.2s;';
                                optionDiv.dataset.value = option.value;

                                const radio = document.createElement('input');
                                radio.type = 'radio';
                                radio.name = 'importMode';
                                radio.value = option.value;
                                radio.checked = index === {defaultIndex};
                                radio.style.cssText = 'margin-top: 2px; cursor: pointer;';

                                const labelDiv = document.createElement('div');
                                labelDiv.style.cssText = 'flex: 1;';

                                const labelText = document.createElement('div');
                                labelText.style.cssText = 'font-weight: 500; color: #333; margin-bottom: 4px;';
                                labelText.textContent = option.label;

                                const descText = document.createElement('div');
                                descText.style.cssText = 'font-size: 13px; color: #666;';
                                descText.textContent = option.desc;

                                labelDiv.appendChild(labelText);
                                labelDiv.appendChild(descText);

                                optionDiv.appendChild(radio);
                                optionDiv.appendChild(labelDiv);

                                // 鼠标悬停效果
                                optionDiv.onmouseover = function() {{
                                    optionDiv.style.borderColor = '#0078d4';
                                    optionDiv.style.backgroundColor = '#f0f8ff';
                                }};
                                optionDiv.onmouseout = function() {{
                                    optionDiv.style.borderColor = '#e0e0e0';
                                    optionDiv.style.backgroundColor = 'white';
                                }};
                                optionDiv.onclick = function() {{
                                    radio.checked = true;
                                    // 移除其他选项的高亮
                                    optionsContainer.querySelectorAll('div[data-value]').forEach(div => {{
                                        div.style.borderColor = '#e0e0e0';
                                        div.style.backgroundColor = 'white';
                                    }});
                                    // 高亮当前选项
                                    optionDiv.style.borderColor = '#0078d4';
                                    optionDiv.style.backgroundColor = '#f0f8ff';
                                }};

                                optionsContainer.appendChild(optionDiv);
                            }});

                            content.appendChild(optionsContainer);

                            // 按钮容器
                            const buttonContainer = document.createElement('div');
                            buttonContainer.style.cssText = 'display: flex; gap: 12px; justify-content: flex-end;';

                            // 取消按钮
                            const cancelButton = document.createElement('button');
                            cancelButton.textContent = '取消';
                            cancelButton.style.cssText = 'padding: 10px 24px; background: #6c757d; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 500;';
                            cancelButton.onmouseover = function() {{ cancelButton.style.background = '#5a6268'; }};
                            cancelButton.onmouseout = function() {{ cancelButton.style.background = '#6c757d'; }};
                            cancelButton.onclick = function() {{
                                document.body.removeChild(dialog);
                                // 返回 null 表示取消
                                window.__xafImportModeResult = null;
                            }};

                            // 确定按钮
                            const okButton = document.createElement('button');
                            okButton.textContent = '确定';
                            okButton.style.cssText = 'padding: 10px 24px; background: #0078d4; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 500;';
                            okButton.onmouseover = function() {{ okButton.style.background = '#106ebe'; }};
                            okButton.onmouseout = function() {{ okButton.style.background = '#0078d4'; }};
                            okButton.onclick = function() {{
                                const selected = optionsContainer.querySelector('input[name=""importMode""]:checked');
                                if (selected) {{
                                    var modeValue = parseInt(selected.value);
                                    console.log('[XAF导入] 选中的导入模式值:', modeValue);
                                    console.log('[XAF导入] 选中的导入模式:', options[modeValue].label);
                                    document.body.removeChild(dialog);
                                    window.__xafImportModeResult = modeValue;
                                }} else {{
                                    alert('请选择一个导入模式');
                                }}
                            }};

                            buttonContainer.appendChild(cancelButton);
                            buttonContainer.appendChild(okButton);
                            content.appendChild(buttonContainer);
                            dialog.appendChild(content);
                            document.body.appendChild(dialog);

                            // 点击背景关闭
                            dialog.onclick = function(e) {{
                                if (e.target === dialog) {{
                                    document.body.removeChild(dialog);
                                    window.__xafImportModeResult = null;
                                }}
                            }};

                            // ESC 键关闭
                            const escapeHandler = function(e) {{
                                if (e.key === 'Escape') {{
                                    document.removeEventListener('keydown', escapeHandler);
                                    const currentDialog = document.getElementById('xaf-import-options-dialog');
                                    if (currentDialog) {{
                                        document.body.removeChild(currentDialog);
                                        window.__xafImportModeResult = null;
                                    }}
                                }}
                            }};
                            document.addEventListener('keydown', escapeHandler);

                            // 等待用户选择
                            return new Promise((resolve) => {{
                                const checkResult = setInterval(function() {{
                                    if (window.__xafImportModeResult !== undefined) {{
                                        clearInterval(checkResult);
                                        const result = window.__xafImportModeResult;
                                        delete window.__xafImportModeResult;
                                        resolve(result);
                                    }}
                                }}, 100);
                            }});
                        }} catch (error) {{
                            console.error('显示导入选项对话框失败:', error);
                            return {defaultIndex}; // 降级到默认值
                        }}
                    }})();
                ";

                var result = await _jsRuntime.InvokeAsync<int?>("eval", script);
                if (result.HasValue)
                {
                    System.Diagnostics.Debug.WriteLine($"[BlazorPlatformFileService] 用户选择的导入模式值: {result.Value} ({(ExcelImportMode)result.Value})");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[BlazorPlatformFileService] 用户取消了导入模式选择");
                }
                return result.HasValue ? (ExcelImportMode)result.Value : (ExcelImportMode?)null;
            }
            catch (Exception ex)
            {
                // 降级: 返回默认模式
                return defaultMode ?? ExcelImportMode.CreateOrUpdate;
            }
#else
            // .NET Framework 版本降级
            return defaultMode ?? ExcelImportMode.CreateOrUpdate;
#endif
        }

        /// <summary>
        /// 导航到指定页面
        /// </summary>
        public async Task<NavigationResult> NavigateToAsync(string url, bool newWindow = false)
        {
            var result = new NavigationResult();

#if NET8_0_OR_GREATER
            try
            {
                if (newWindow)
                {
                    // 在新窗口中打开
                    if (_jsRuntime != null)
                    {
                        var safeUrl = EscapeJavaScriptString(url);
                        await _jsRuntime.InvokeVoidAsync("open", safeUrl, "_blank");
                    }
                    else
                    {
                        result.IsSuccess = false;
                        result.ErrorMessage = "JavaScript运行时不可用";
                        return result;
                    }
                }
                else
                {
                    // 在当前窗口中导航
                    if (_navigationManager != null)
                    {
                        _navigationManager.NavigateTo(url);
                    }
                    else
                    {
                        result.IsSuccess = false;
                        result.ErrorMessage = "导航管理器不可用";
                        return result;
                    }
                }
                
                result.IsSuccess = true;
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.ErrorMessage = $"导航失败: {ex.Message}";
            }
#else
            result.IsSuccess = false;
            result.ErrorMessage = "当前.NET Framework版本不支持Blazor导航";
#endif

            return result;
        }

        /// <summary>
        /// 转义JavaScript字符串
        /// </summary>
        private string EscapeJavaScriptString(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            return input
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("'", "\\'")
                .Replace("\n", "\\n")
                .Replace("\r", "\\r")
                .Replace("\t", "\\t")
                .Replace("<", "\\u003c")
                .Replace(">", "\\u003e");
        }
    }
}