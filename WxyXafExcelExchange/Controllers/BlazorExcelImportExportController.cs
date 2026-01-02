using System;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using DevExpress.ExpressApp;
using Wxy.Xaf.ExcelExchange.Services;
using System.ComponentModel;
using System.Diagnostics;

#if NET8_0_OR_GREATER
using Microsoft.JSInterop;
using Microsoft.AspNetCore.Components;
#endif

namespace Wxy.Xaf.ExcelExchange.Controllers
{
    /// <summary>
    /// Blazor平台Excel导入导出控制器
    /// </summary>
    [ToolboxItemFilter("Xaf.Platform.Blazor")]
    public class BlazorExcelImportExportController : ExcelImportExportControllerBase
    {
        private static bool? _isBlazorPlatform = null;

        public BlazorExcelImportExportController()
        {
            // 检查平台，如果不是Blazor平台则禁用控制器
            if (!IsBlazorPlatform())
            {
                this.Active["PlatformCheck"] = false;
            }
        }

        protected override void OnActivated()
        {
            base.OnActivated();
        }

        protected override void OnViewChanged()
        {
            base.OnViewChanged();
        }

        /// <summary>
        /// 检查是否为Blazor平台（静态缓存结果）
        /// </summary>
        private static bool IsBlazorPlatform()
        {
            if (_isBlazorPlatform.HasValue)
            {
                return _isBlazorPlatform.Value;
            }

            try
            {
                // 方法1: 检查是否存在Blazor相关的类型
                var blazorType = Type.GetType("Microsoft.AspNetCore.Components.ComponentBase, Microsoft.AspNetCore.Components");
                bool hasBlazorType = blazorType != null;

                // 方法2: 检查当前应用程序域是否加载了Blazor程序集
                var blazorAssembly = AppDomain.CurrentDomain.GetAssemblies()
                    .FirstOrDefault(a => a.GetName().Name?.Contains("Blazor") == true ||
                                        a.GetName().Name?.Contains("AspNetCore") == true);
                bool hasBlazorAssembly = blazorAssembly != null;

                // 方法3: 检查是否存在WinForms相关的类型（排除法）
                var winFormsType = Type.GetType("System.Windows.Forms.Application, System.Windows.Forms");
                bool hasWinFormsType = winFormsType != null;

                // 方法4: 检查应用程序类型名称
                var entryAssembly = Assembly.GetEntryAssembly();
                bool hasBlazorEntryPoint = false;

                if (entryAssembly != null)
                {
                    var entryAssemblyName = entryAssembly.GetName().Name;
                    // More specific Blazor detection - look for common Blazor naming patterns
                    hasBlazorEntryPoint = entryAssemblyName?.Contains("Blazor") == true ||
                                         entryAssemblyName?.Contains("Server") == true ||
                                         entryAssemblyName?.Contains("blazor") == true ||
                                         entryAssemblyName?.EndsWith(".Server") == true;
                }


                // 严格判断：必须有Blazor类型 AND 有Blazor程序集 AND (没有WinForms类型 OR 有明确的Blazor入口点)
                // 这样可以避免在WinForms应用中误判为Blazor平台
                bool result = hasBlazorType && hasBlazorAssembly && (!hasWinFormsType || hasBlazorEntryPoint);

                _isBlazorPlatform = result;
                return result;
            }
            catch (Exception ex)
            {
                // 如果检测失败，默认不允许Blazor控制器实例化（保守策略）
                Debug.WriteLine($"[BlazorController] 平台检测异常: {ex.Message}");
                _isBlazorPlatform = false;
                return false;
            }
        }

        /// <summary>
        /// 获取操作ID前缀
        /// </summary>
        protected override string GetActionIdPrefix()
        {
            return "Blazor";
        }

        /// <summary>
        /// 获取Excel服务
        /// </summary>
        protected override IExcelImportExportService GetExcelService()
        {
            // 尝试从服务提供者获取
            if (Application?.ServiceProvider != null)
            {
                var service = Application.ServiceProvider.GetService(typeof(IExcelImportExportService)) as IExcelImportExportService;
                if (service != null)
                {
                    return service;
                }
            }

            // 降级到直接创建实例
            return new ExcelImportExportService();
        }

        /// <summary>
        /// 获取平台文件服务
        /// </summary>
        protected override IPlatformFileService GetPlatformFileService()
        {
            // 尝试从服务提供者获取
            if (Application?.ServiceProvider != null)
            {
                var service = Application.ServiceProvider.GetService(typeof(IPlatformFileService)) as IPlatformFileService;
                if (service != null)
                {
                    return service;
                }
            }

#if NET8_0_OR_GREATER
            // 尝试获取Blazor特定的服务
            try
            {
                if (Application?.ServiceProvider != null)
                {
                    var jsRuntime = Application.ServiceProvider.GetService(typeof(IJSRuntime)) as IJSRuntime;
                    var navigationManager = Application.ServiceProvider.GetService(typeof(NavigationManager)) as NavigationManager;

                    if (jsRuntime != null && navigationManager != null && Application != null)
                    {
                        return new BlazorPlatformFileService(Application, jsRuntime, navigationManager);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"创建Blazor平台服务失败: {ex.Message}", ex);
            }
#endif

            // 直接创建基础实例，允许Application为null
            // 如果Application为null，某些功能可能受限，但不会导致崩溃
            return new BlazorPlatformFileService(Application);
        }

        /// <summary>
        /// 执行平台特定的导入流程
        /// </summary>
        protected override async Task ExecutePlatformSpecificImport(Type objectType)
        {
            try
            {
                LogDebug("Blazor: 开始导入流程");
                LogDebug($"Blazor: 导入对象类型: {objectType.FullName}");

                // 在Blazor中，直接使用文件上传对话框
                var fileResult = await PlatformFileService.PickFileAsync(new[] { ".xlsx", ".xls", ".csv" });

                if (!fileResult.IsSuccess || fileResult.FileContent == null || fileResult.FileContent.Length == 0)
                {
                    LogDebug("Blazor: 文件选择被取消或文件为空");
                    return;
                }

                LogDebug($"Blazor: 已选择文件: {fileResult.FileName}, 大小: {fileResult.FileContent.Length} 字节");

                // **新增**: 显示导入模式选择对话框
                ExcelImportMode? selectedMode = await PlatformFileService.ShowImportOptionsDialogAsync(
                    $"选择 {objectType.Name} 的导入模式",
                    GetDefaultImportMode(objectType)
                );

                // 如果用户取消,则中止导入
                if (!selectedMode.HasValue)
                {
                    LogDebug("Blazor: 用户取消了导入模式选择");
                    return;
                }

                LogDebug($"Blazor: 用户选择的导入模式: {selectedMode.Value} (枚举值: {(int)selectedMode.Value})");
                LogDebug($"Blazor: 模式说明 - CreateOnly=0, UpdateOnly=1, CreateOrUpdate=2, ReplaceAll=3");

                // 使用基类的核心导入逻辑处理文件,传入用户选择的模式
                var importResult = await ProcessFileImport(
                    fileResult.FileContent,
                    fileResult.FileName,
                    objectType,
                    selectedMode.Value  // 传入用户选择的模式
                );

                // 显示导入结果
                await ShowImportResult(importResult);

                LogDebug($"Blazor: 导入流程完成 - 成功: {importResult.SuccessCount}, 失败: {importResult.FailureCount}");
            }
            catch (Exception ex)
            {
                LogError($"Blazor导入流程异常: {ex.Message}", ex);
                await PlatformFileService.ShowMessageAsync(
                    $"导入失败: {ex.Message}",
                    "导入异常",
                    MessageType.Error);
            }
        }

        /// <summary>
        /// 获取默认导入模式
        /// </summary>
        private ExcelImportMode GetDefaultImportMode(Type objectType)
        {
            try
            {
                var config = ConfigurationManager.GetConfiguration(objectType, View?.ObjectTypeInfo);
                var classConfig = config.ClassConfiguration;

                // 从类配置中读取默认模式
                if (classConfig != null)
                {
                    // 这里可以读取 ImportDuplicateStrategy,但为了简化,我们返回默认值
                    // 实际使用中,可以从 ExcelImportExportAttribute 读取配置
                }

                return ExcelImportMode.CreateOrUpdate; // 默认值
            }
            catch
            {
                return ExcelImportMode.CreateOrUpdate; // 降级默认值
            }
        }

        /// <summary>
        /// 显示手动导航指令
        /// </summary>
        private async Task ShowManualNavigationInstructions(Type objectType, string uploadUrl)
        {
            try
            {
                var message = $"📁 Excel导入页面导航指南\n\n" +
                             $"您正在尝试导入 {objectType.Name} 数据。\n" +
                             $"请手动导航到Excel上传页面。\n\n" +
                             $"📋 操作步骤：\n\n" +
                             $"方法1 - 直接地址导航\n" +
                             $"   1. 复制以下URL：{uploadUrl}\n" +
                             $"   2. 粘贴到浏览器地址栏\n" +
                             $"   3. 按Enter键访问\n\n" +
                             $"方法2 - 新标签页导航\n" +
                             $"   1. 按 Ctrl+T 打开新标签页\n" +
                             $"   2. 复制粘贴上述URL\n" +
                             $"   3. 按Enter键访问\n\n" +
                             $"💡 提示：\n" +
                             $"- 确保您的浏览器允许JavaScript执行\n" +
                             $"- 如果问题持续存在，请联系系统管理员";

                await PlatformFileService.ShowMessageAsync(message, "导航指南", MessageType.Information);
            }
            catch (Exception ex)
            {
                LogError($"显示导航指令失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 处理Blazor文件导入（供外部调用）
        /// </summary>
        public async Task<ExcelImportResult> ProcessBlazorFileImportAsync(byte[] fileContent, string fileName, Type objectType)
        {
            try
            {
                LogDebug($"Blazor: 处理文件导入 - {fileName}");

                // 使用基类的核心导入逻辑
                var result = await ProcessFileImport(fileContent, fileName, objectType);

                LogDebug($"Blazor: 文件导入完成 - 成功: {result.SuccessCount}, 失败: {result.FailureCount}");

                return result;
            }
            catch (Exception ex)
            {
                LogError($"Blazor文件导入处理异常: {ex.Message}", ex);

                var result = new ExcelImportResult
                {
                    IsSuccess = false,
                    ErrorMessage = ex.Message
                };
                result.Errors.Add(new ExcelImportError
                {
                    RowNumber = 0,
                    FieldName = "系统",
                    ErrorMessage = ex.Message,
                    ErrorType = ExcelImportErrorType.SystemError
                });
                return result;
            }
        }

        /// <summary>
        /// 显示导入结果（Blazor特定实现）
        /// </summary>
        protected override async Task ShowImportResult(ExcelImportResult result)
        {
            try
            {
                // 构建完整的导入结果信息
                var message = BuildImportResultMessage(result);

                // 根据结果显示不同类型的消息
                MessageType messageType;
                string title;

                if (result.IsSuccess)
                {
                    if (result.FailureCount == 0 && result.Warnings.Count == 0)
                    {
                        title = "导入完成";
                        messageType = MessageType.Success;
                    }
                    else if (result.FailureCount > 0)
                    {
                        title = "导入部分完成";
                        messageType = MessageType.Warning;
                    }
                    else
                    {
                        title = "导入完成";
                        messageType = MessageType.Success;
                    }
                }
                else
                {
                    title = "导入失败";
                    messageType = MessageType.Error;
                }

                await PlatformFileService.ShowMessageAsync(message, title, messageType);

                // **Blazor 特殊处理**: 导入成功后提示用户刷新页面
                // 这是因为 Blazor Grid 在刷新时可能会遇到 XPO 跨线程验证问题
                if (result.IsSuccess && result.SuccessCount > 0)
                {
                    await System.Threading.Tasks.Task.Delay(500);
                    await PlatformFileService.ShowMessageAsync(
                        "✓ 数据已成功导入到数据库\n\n💡 提示: 如果列表未显示新数据,请按 F5 刷新页面",
                        "导入提示",
                        MessageType.Information);
                }
            }
            catch (Exception ex)
            {
                LogError($"显示导入结果失败: {ex.Message}", ex);
                await PlatformFileService.ShowMessageAsync(
                    $"导入结果显示异常: {ex.Message}",
                    "显示异常",
                    MessageType.Error);
            }
        }

        /// <summary>
        /// 构建导入结果消息
        /// </summary>
        private string BuildImportResultMessage(ExcelImportResult result)
        {
            var message = new System.Text.StringBuilder();

            if (result.IsSuccess)
            {
                // 判断是否完全成功
                if (result.FailureCount == 0 && result.Warnings.Count == 0)
                {
                    message.AppendLine("✓ 数据导入成功");
                    message.AppendLine();
                }
                else if (result.FailureCount > 0)
                {
                    message.AppendLine("△ 导入部分完成");
                    message.AppendLine();
                }
                else
                {
                    message.AppendLine("✓ 数据导入成功");
                    message.AppendLine();
                }
            }
            else
            {
                message.AppendLine("✗ 导入失败");
                message.AppendLine();
            }

            // 统计信息
            message.AppendLine("═══════════════════════════════════════════════════════════════");
            message.AppendLine("统计信息");
            message.AppendLine("═══════════════════════════════════════════════════════════════");
            message.AppendLine($"总记录数: {result.TotalRecords:N0}");
            message.AppendLine($"成功导入: {result.SuccessCount:N0}");

            if (result.FailureCount > 0)
            {
                message.AppendLine($"导入失败: {result.FailureCount:N0}");
            }

            if (result.Warnings.Count > 0)
            {
                message.AppendLine($"警告数量: {result.Warnings.Count:N0}");
            }

            message.AppendLine();

            // 警告详情
            if (result.Warnings.Count > 0)
            {
                message.AppendLine("─────────────────────────────────────────────────────────────────");
                message.AppendLine("警告详情:");
                message.AppendLine("─────────────────────────────────────────────────────────────────");
                message.AppendLine();

                for (int i = 0; i < result.Warnings.Count; i++)
                {
                    var warning = result.Warnings[i];
                    message.AppendLine($"[{i + 1}] 行 {warning.RowNumber} - {warning.FieldName}");
                    message.AppendLine($"    {warning.WarningMessage}");
                    message.AppendLine();
                }
            }

            // 错误详情
            if (result.Errors.Count > 0)
            {
                message.AppendLine("═══════════════════════════════════════════════════════════════");
                message.AppendLine("失败记录详情:");
                message.AppendLine("═══════════════════════════════════════════════════════════════");
                message.AppendLine();

                for (int i = 0; i < result.Errors.Count; i++)
                {
                    var error = result.Errors[i];
                    message.AppendLine($"[{i + 1}] 行 {error.RowNumber} - {error.FieldName}");
                    message.AppendLine($"    错误: {error.ErrorMessage}");

                    if (!string.IsNullOrEmpty(error.OriginalValue))
                    {
                        message.AppendLine($"    原始值: {error.OriginalValue}");
                    }
                    message.AppendLine();
                }
            }

            if (result.IsSuccess)
            {
                message.AppendLine("💡 提示: 已导入成功的记录已保存到数据库");
                message.AppendLine("💡 点击「复制」按钮可复制完整信息");
            }

            return message.ToString();
        }

        /// <summary>
        /// 获取导出选项（Blazor特定）
        /// </summary>
        protected override ExcelExportOptions GetExportOptions(Type objectType)
        {
            var baseOptions = base.GetExportOptions(objectType);

            // Blazor平台使用XLSX格式以支持多Sheet功能
            baseOptions.Format = ExcelFormat.Xlsx;

            return baseOptions;
        }

        /// <summary>
        /// 获取导入选项（Blazor特定）
        /// </summary>
        protected override ExcelImportOptions GetImportOptions(Type objectType)
        {
            var baseOptions = base.GetImportOptions(objectType);

            // Blazor平台使用较小的批次，避免长时间阻塞UI
            baseOptions.BatchSize = 500;
            baseOptions.MaxErrors = 50;

            return baseOptions;
        }

        /// <summary>
        /// 获取服务诊断信息（用于调试）
        /// </summary>
        public string GetServiceDiagnosticInfo()
        {
            var info = new System.Text.StringBuilder();

            try
            {
                info.AppendLine("Blazor服务状态诊断：");
                info.AppendLine($"- Application: {(Application != null ? "可用" : "null")}");
                info.AppendLine($"- ServiceProvider: {(Application?.ServiceProvider != null ? "可用" : "null")}");

                if (Application?.ServiceProvider != null)
                {
#if NET8_0_OR_GREATER
                    try
                    {
                        var jsRuntime = Application.ServiceProvider.GetService(typeof(IJSRuntime));
                        info.AppendLine($"- IJSRuntime: {(jsRuntime != null ? "可用" : "不可用")}");

                        var navigationManager = Application.ServiceProvider.GetService(typeof(NavigationManager));
                        info.AppendLine($"- NavigationManager: {(navigationManager != null ? "可用" : "不可用")}");
                    }
                    catch (Exception ex)
                    {
                        info.AppendLine($"- Blazor服务检查失败: {ex.Message}");
                    }
#else
                    info.AppendLine("- 当前运行在.NET Framework模式");
#endif
                }

                info.AppendLine($"- ExcelService: {(ExcelService != null ? "可用" : "不可用")}");
                info.AppendLine($"- PlatformFileService: {(PlatformFileService != null ? "可用" : "不可用")}");
            }
            catch (Exception ex)
            {
                info.AppendLine($"诊断信息收集失败: {ex.Message}");
            }

            return info.ToString();
        }
    }
}
