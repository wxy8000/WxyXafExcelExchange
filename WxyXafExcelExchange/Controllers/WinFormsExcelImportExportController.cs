using System;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using DevExpress.ExpressApp;
using Wxy.Xaf.ExcelExchange.Services;
using System.ComponentModel;

#if NET8_0_WINDOWS || NET462_OR_GREATER

namespace Wxy.Xaf.ExcelExchange.Controllers
{
    /// <summary>
    /// WinForms平台Excel导入导出控制器
    /// 仅在Windows平台编译
    /// </summary>
    [ToolboxItemFilter("Xaf.Platform.Win")]
    public class WinFormsExcelImportExportController : ExcelImportExportControllerBase
    {
        private static bool? _isWinFormsPlatform = null;

        public WinFormsExcelImportExportController()
        {
            // 检查平台，如果不是WinForms平台则禁用控制器
            if (!IsWinFormsPlatform())
            {
                this.Active["PlatformCheck"] = false;
            }
        }

        protected override void OnActivated()
        {
            // 平台检测已在构造函数中完成，这里不需要重复检查
            base.OnActivated();
        }

        /// <summary>
        /// 检查是否为WinForms平台（静态缓存结果）
        /// </summary>
        private static bool IsWinFormsPlatform()
        {
            if (_isWinFormsPlatform.HasValue)
            {
                return _isWinFormsPlatform.Value;
            }

            try
            {
                // 方法1: 检查是否存在WinForms相关的类型
                var winFormsAppType = Type.GetType("System.Windows.Forms.Application, System.Windows.Forms");
                bool hasWinFormsType = winFormsAppType != null;
                
                // 方法2: 检查当前应用程序域是否加载了WinForms程序集
                var winFormsAssembly = AppDomain.CurrentDomain.GetAssemblies()
                    .FirstOrDefault(a => a.GetName().Name == "System.Windows.Forms");
                bool hasWinFormsAssembly = winFormsAssembly != null;
                
                // 方法3: 检查是否存在Blazor特定的类型（排除法）
                var blazorComponentType = Type.GetType("Microsoft.AspNetCore.Components.ComponentBase, Microsoft.AspNetCore.Components");
                bool hasBlazorType = blazorComponentType != null;
                
                // 方法4: 检查应用程序类型名称
                var currentAppDomain = AppDomain.CurrentDomain;
                var entryAssembly = Assembly.GetEntryAssembly();
                bool hasWinFormsEntryPoint = false;
                
                if (entryAssembly != null)
                {
                    var entryAssemblyName = entryAssembly.GetName().Name;
                    // More specific WinForms detection - look for common WinForms naming patterns
                    hasWinFormsEntryPoint = entryAssemblyName?.Contains("Win") == true || 
                                          entryAssemblyName?.Contains("WinForms") == true ||
                                          entryAssemblyName?.Contains("winform") == true ||
                                          entryAssemblyName?.EndsWith(".Win") == true;
                }
                
                // 更宽松的判断逻辑：
                // 1. 必须有WinForms类型和程序集
                // 2. 如果有Blazor类型，则需要额外的WinForms证据（如入口程序集名称）
                // 3. 如果没有Blazor类型，则认为是WinForms平台
                bool result;
                if (hasWinFormsType && hasWinFormsAssembly)
                {
                    if (hasBlazorType)
                    {
                        // 有Blazor类型时需要更多证据证明是WinForms
                        result = hasWinFormsEntryPoint;
                    }
                    else
                    {
                        // 没有Blazor类型，有WinForms类型和程序集就认为是WinForms
                        result = true;
                    }
                }
                else
                {
                    result = false;
                }
                
                                
                _isWinFormsPlatform = result;
                return result;
            }
            catch
            {
                                // 如果检测失败，默认允许WinForms控制器实例化（保守策略）
                _isWinFormsPlatform = true;
                return true;
            }
        }
        /// <summary>
        /// 获取操作ID前缀
        /// </summary>
        protected override string GetActionIdPrefix()
        {
            return "WinForms";
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

            // 直接创建实例，允许Application为null（WinFormsPlatformFileService已支持）
            // 如果Application为null，某些功能可能受限，但不会导致崩溃
            return new WinFormsPlatformFileService(Application);
        }

        /// <summary>
        /// 执行平台特定的导入流程
        /// </summary>
        protected override async Task ExecutePlatformSpecificImport(Type objectType)
        {
            try
            {
                // 显示文件选择对话框
                var fileSelectionOptions = new FileSelectionOptions
                {
                    Filter = "Excel文件 (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|CSV文件 (*.csv)|*.csv|Excel 2007+ (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls|所有文件 (*.*)|*.*",
                    Title = $"选择要导入的{objectType.Name}文件",
                    Multiselect = false
                };

                var selectionResult = await PlatformFileService.ShowFileSelectionDialogAsync(fileSelectionOptions);

                if (!selectionResult.IsSuccess)
                {
                    if (selectionResult.IsCancelled)
                    {
                        return;
                    }
                    else
                    {
                        await PlatformFileService.ShowMessageAsync(
                            $"文件选择失败: {selectionResult.ErrorMessage}",
                            "导入失败",
                            MessageType.Error);
                        return;
                    }
                }

                if (selectionResult.FileContents == null || selectionResult.FileContents.Length == 0)
                {
                    await PlatformFileService.ShowMessageAsync(
                        "未选择任何文件",
                        "导入失败",
                        MessageType.Warning);
                    return;
                }

                var fileContent = selectionResult.FileContents[0];

                // **新增**: 显示导入模式选择对话框
                ExcelImportMode? selectedMode = await PlatformFileService.ShowImportOptionsDialogAsync(
                    $"选择 {objectType.Name} 的导入模式",
                    GetDefaultImportMode(objectType)
                );

                // 如果用户取消,则中止导入
                if (!selectedMode.HasValue)
                {
                    return;
                }

                LogDebug($"WinForms: 用户选择的导入模式: {selectedMode.Value}");

                // **关键修改**: 直接在当前上下文执行导入,不使用额外的线程调度
                // ProcessFileImport 内部已经正确处理了 ObjectSpace 的线程安全性
                var importResult = await ProcessFileImport(
                    fileContent.Content,
                    fileContent.FileName,
                    objectType,
                    selectedMode.Value
                );

                // 显示导入结果
                await ShowImportResult(importResult);
            }
            catch (Exception ex)
            {
                LogError($"WinForms导入流程异常: {ex.Message}", ex);
                await PlatformFileService.ShowMessageAsync(
                    $"导入过程中发生异常: {ex.Message}",
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
        /// 显示导入结果（WinForms特定实现）
        /// </summary>
        protected override async Task ShowImportResult(ExcelImportResult result)
        {
            try
            {
                if (result.IsSuccess)
                {
                    // 使用自定义对话框显示导入结果
                    await ShowImportResultDialog(result);

                    // 刷新视图以显示新数据 (在UI线程执行)
                    if (View != null)
                    {
                        View.ObjectSpace.CommitChanges();
                        // ObjectSpace.Refresh() 可能导致多线程问题,使用 CommitChanges 代替
                    }
                }
                else
                {
                    var message = $"❌ 导入失败！\n\n" +
                                 $"错误信息: {result.ErrorMessage}";

                    if (result.Errors.Count > 0)
                    {
                        message += $"\n\n📋 详细错误信息 (共{result.Errors.Count}条):\n";

                        for (int i = 0; i < result.Errors.Count; i++)
                        {
                            var error = result.Errors[i];
                            message += $"\n{i + 1}. 行{error.RowNumber}: {error.FieldName}\n";
                            message += $"   {error.ErrorMessage}";

                            if (!string.IsNullOrEmpty(error.OriginalValue))
                            {
                                message += $"\n   原始值: {error.OriginalValue}";
                            }
                        }
                    }

                    await PlatformFileService.ShowMessageAsync(message, "导入失败", MessageType.Error);
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
        /// 显示自定义导入结果对话框
        /// </summary>
        private async Task ShowImportResultDialog(ExcelImportResult result)
        {
            await Wxy.Xaf.ExcelExchange.Threading.UIThreadDispatcher.InvokeOnUIThreadAsync(() =>
            {
#if NET8_0_WINDOWS || NET462_OR_GREATER
                using (var form = new System.Windows.Forms.Form())
                {
                    form.Text = "导入完成";
                    form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                    form.MaximizeBox = true;
                    form.MinimizeBox = true;
                    form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                    form.ClientSize = new System.Drawing.Size(700, 550);
                    form.BackColor = System.Drawing.Color.White;
                    form.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F);
                    form.MinimumSize = new System.Drawing.Size(600, 400);

                    // 标题
                    var titleLabel = new System.Windows.Forms.Label
                    {
                        Text = result.FailureCount == 0 && result.Warnings.Count == 0
                            ? "✓ 数据导入成功"
                            : "△ 导入完成",
                        Font = new System.Drawing.Font("Microsoft YaHei UI", 14F, System.Drawing.FontStyle.Bold),
                        Location = new System.Drawing.Point(20, 20),
                        Size = new System.Drawing.Size(660, 35),
                        ForeColor = result.FailureCount == 0 && result.Warnings.Count == 0
                            ? System.Drawing.Color.FromArgb(0, 122, 102)
                            : System.Drawing.Color.FromArgb(255, 152, 0)
                    };
                    form.Controls.Add(titleLabel);

                    // 详情文本框 (放在前面,以便按钮可以引用)
                    var detailsTextBox = new System.Windows.Forms.TextBox
                    {
                        Multiline = true,
                        ReadOnly = true,
                        ScrollBars = System.Windows.Forms.ScrollBars.Vertical,
                        Location = new System.Drawing.Point(20, 110),
                        Size = new System.Drawing.Size(660, 370),
                        Font = new System.Drawing.Font("Consolas", 9F),
                        BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                        BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                        Text = BuildDetailsText(result)
                    };
                    form.Controls.Add(detailsTextBox);

                    // 按钮区域
                    var copyButton = new System.Windows.Forms.Button
                    {
                        Text = "复制到剪贴板",
                        Location = new System.Drawing.Point(20, 65),
                        Size = new System.Drawing.Size(120, 30),
                        BackColor = System.Drawing.Color.FromArgb(0, 122, 102),
                        ForeColor = System.Drawing.Color.White,
                        FlatStyle = System.Windows.Forms.FlatStyle.Flat,
                        Font = new System.Drawing.Font("Microsoft YaHei UI", 9F)
                    };
                    copyButton.FlatAppearance.BorderSize = 0;
                    copyButton.Click += (s, e) =>
                    {
                        System.Windows.Forms.Clipboard.SetText(detailsTextBox.Text);
                        copyButton.Text = "✓ 已复制";
                        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                        timer.Interval = 1500;
                        timer.Tick += (timerSender, timerE) =>
                        {
                            copyButton.Text = "复制到剪贴板";
                            timer.Stop();
                            ((System.Windows.Forms.Timer)timerSender).Dispose();
                        };
                        timer.Start();
                    };
                    form.Controls.Add(copyButton);

                    var okButton = new System.Windows.Forms.Button
                    {
                        Text = "确定",
                        DialogResult = System.Windows.Forms.DialogResult.OK,
                        Location = new System.Drawing.Point(600, 65),
                        Size = new System.Drawing.Size(80, 30),
                        BackColor = System.Drawing.Color.FromArgb(0, 122, 102),
                        ForeColor = System.Drawing.Color.White,
                        FlatStyle = System.Windows.Forms.FlatStyle.Flat,
                        Font = new System.Drawing.Font("Microsoft YaHei UI", 9F)
                    };
                    okButton.FlatAppearance.BorderSize = 0;
                    form.Controls.Add(okButton);
                    form.AcceptButton = okButton;

                    form.ShowDialog();
#endif
                }
            });
        }

#if NET8_0_WINDOWS || NET462_OR_GREATER
        /// <summary>
        /// 添加统计行
        /// </summary>
        private void AddStatRow(System.Windows.Forms.Form form, string label, string value, ref int yPos, System.Drawing.Color? valueColor = null)
        {
            var lbl = new System.Windows.Forms.Label
            {
                Text = label,
                Location = new System.Drawing.Point(30, yPos),
                Size = new System.Drawing.Size(150, 25),
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(102, 102, 102)
            };
            form.Controls.Add(lbl);

            var val = new System.Windows.Forms.Label
            {
                Text = value,
                Location = new System.Drawing.Point(180, yPos),
                Size = new System.Drawing.Size(390, 25),
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F),
                ForeColor = valueColor ?? System.Drawing.Color.FromArgb(51, 51, 51)
            };
            form.Controls.Add(val);

            yPos += 30;
        }

        /// <summary>
        /// 构建详细信息文本
        /// </summary>
        private string BuildDetailsText(ExcelImportResult result)
        {
            var details = new System.Text.StringBuilder();

            details.AppendLine("═══════════════════════════════════════════════════════════════");
            details.AppendLine("导入结果详情");
            details.AppendLine("═══════════════════════════════════════════════════════════════\n");

            details.AppendLine($"总记录数: {result.TotalRecords:N0}");
            details.AppendLine($"成功导入: {result.SuccessCount:N0}");
            details.AppendLine($"导入失败: {result.FailureCount:N0}");
            details.AppendLine($"警告数量: {result.Warnings.Count:N0}");
            details.AppendLine();

            // 警告详情
            if (result.Warnings.Count > 0)
            {
                details.AppendLine("─────────────────────────────────────────────────────────────────");
                details.AppendLine("警告详情:");
                details.AppendLine("─────────────────────────────────────────────────────────────────\n");

                for (int i = 0; i < Math.Min(result.Warnings.Count, 50); i++) // 最多显示50条
                {
                    var warning = result.Warnings[i];
                    details.AppendLine($"[{i + 1}] 行 {warning.RowNumber} - {warning.FieldName}");
                    details.AppendLine($"    {warning.WarningMessage}");
                    details.AppendLine();
                }

                if (result.Warnings.Count > 50)
                {
                    details.AppendLine($"... 还有 {result.Warnings.Count - 50} 条警告未显示");
                    details.AppendLine();
                }
            }

            // 错误详情
            if (result.Errors.Count > 0)
            {
                details.AppendLine("═══════════════════════════════════════════════════════════════");
                details.AppendLine("失败记录详情:");
                details.AppendLine("═══════════════════════════════════════════════════════════════\n");

                for (int i = 0; i < Math.Min(result.Errors.Count, 50); i++) // 最多显示50条
                {
                    var error = result.Errors[i];
                    details.AppendLine($"[{i + 1}] 行 {error.RowNumber} - {error.FieldName}");
                    details.AppendLine($"    错误: {error.ErrorMessage}");

                    if (!string.IsNullOrEmpty(error.OriginalValue))
                    {
                        details.AppendLine($"    原始值: {error.OriginalValue}");
                    }
                    details.AppendLine();
                }

                if (result.Errors.Count > 50)
                {
                    details.AppendLine($"... 还有 {result.Errors.Count - 50} 条错误未显示");
                }
            }

            return details.ToString();
        }
#endif

        /// <summary>
        /// 获取导出选项（WinForms特定）
        /// </summary>
        protected override ExcelExportOptions GetExportOptions(Type objectType)
        {
            var baseOptions = base.GetExportOptions(objectType);

            // WinForms平台使用XLSX格式以支持多Sheet功能
            baseOptions.Format = ExcelFormat.Xlsx;

            return baseOptions;
        }

        /// <summary>
        /// 获取导入选项（WinForms特定）
        /// </summary>
        protected override ExcelImportOptions GetImportOptions(Type objectType)
        {
            var baseOptions = base.GetImportOptions(objectType);
            
            // WinForms平台可以处理更大的批次
            baseOptions.BatchSize = 2000;
            baseOptions.MaxErrors = 200;

            return baseOptions;
        }
    }
}

#endif // NET8_0_WINDOWS || NET462_OR_GREATER
