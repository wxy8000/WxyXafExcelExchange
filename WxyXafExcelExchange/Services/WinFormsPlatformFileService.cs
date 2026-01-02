using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DevExpress.ExpressApp;
using Wxy.Xaf.ExcelExchange.Threading;

#if NET8_0_WINDOWS || NET462_OR_GREATER
using System.Windows.Forms;

namespace Wxy.Xaf.ExcelExchange.Services
{
    /// <summary>
    /// WinForms平台文件服务实现
    /// 仅在Windows平台编译
    /// </summary>
    public class WinFormsPlatformFileService : IPlatformFileService
    {
        private readonly XafApplication _application;
        public WinFormsPlatformFileService(XafApplication application)
        {
            _application = application; // 允许为null，在使用时再检查
        }
        /// <summary>
        /// 显示文件选择对话框
        /// </summary>
        public async Task<FileSelectionResult> ShowFileSelectionDialogAsync(FileSelectionOptions options)
        {
            try
            {
                // 使用UI线程调度器确保在正确的线程上执行
                return await UIThreadDispatcher.InvokeOnUIThreadAsync(() =>
                {
                    return ShowFileSelectionDialogSync(options);
                });
            }
            catch (STAThreadException staEx)
            {
                
                var result = new FileSelectionResult
                {
                    IsSuccess = false,
                    ErrorMessage = $"文件选择失败: {staEx.Message}"
                };
                return result;
            }
            catch (Exception ex)
            {
                
                // 检查是否可能是STA线程问题
                if (STAThreadException.IsPotentialSTAThreadIssue(ex))
                {
                    var staEx = STAThreadException.CreateForFileDialog("文件选择", ex);
                    
                    var result = new FileSelectionResult
                    {
                        IsSuccess = false,
                        ErrorMessage = $"文件选择失败: {staEx.Message}"
                    };
                    return result;
                }
                
                var errorResult = new FileSelectionResult
                {
                    IsSuccess = false,
                    ErrorMessage = $"文件选择失败: {ex.Message}"
                };
                return errorResult;
            }
        }
        /// <summary>
        /// 同步显示文件选择对话框（必须在STA线程调用）
        /// </summary>
        private FileSelectionResult ShowFileSelectionDialogSync(FileSelectionOptions options)
        {
            var result = new FileSelectionResult();
            try
            {
                using (var openDialog = new OpenFileDialog())
                {
                    openDialog.Filter = options.Filter;
                    openDialog.Title = options.Title;
                    openDialog.Multiselect = options.Multiselect;
                    openDialog.CheckFileExists = true;
                    openDialog.CheckPathExists = true;
                    if (!string.IsNullOrEmpty(options.InitialDirectory))
                    {
                        openDialog.InitialDirectory = options.InitialDirectory;
                    }
                    if (!string.IsNullOrEmpty(options.DefaultFileName))
                    {
                        openDialog.FileName = options.DefaultFileName;
                    }
                    var dialogResult = openDialog.ShowDialog();
                    if (dialogResult == DialogResult.OK)
                    {
                        result.IsSuccess = true;
                        result.FilePaths = openDialog.FileNames;
                        // 读取文件内容
                        var fileContents = new FileContentInfo[openDialog.FileNames.Length];
                        for (int i = 0; i < openDialog.FileNames.Length; i++)
                        {
                            var filePath = openDialog.FileNames[i];
                            var fileInfo = new FileInfo(filePath);
                            
                            fileContents[i] = new FileContentInfo
                            {
                                FileName = fileInfo.Name,
                                Content = File.ReadAllBytes(filePath),
                                Size = fileInfo.Length,
                                MimeType = GetMimeType(fileInfo.Extension)
                            };
                        }
                        result.FileContents = fileContents;
                    }
                    else
                    {
                        result.IsCancelled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.ErrorMessage = $"文件选择失败: {ex.Message}";
            }
            return result;
        }
        /// <summary>
        /// 选择文件（简化的便捷方法）
        /// </summary>
        public async Task<FileSelectionResult> PickFileAsync(string[] allowedExtensions = null)
        {
            // 构建文件过滤器
            string filter = "所有文件 (*.*)|*.*";
            if (allowedExtensions != null && allowedExtensions.Length > 0)
            {
                var filters = allowedExtensions.Select(ext =>
                {
                    var desc = ext.ToUpperInvariant() switch
                    {
                        ".XLSX" => "Excel工作簿",
                        ".XLS" => "Excel工作簿",
                        ".CSV" => "CSV文件",
                        _ => "文件"
                    };
                    return $"{desc} ({ext})|{ext}";
                });
                filter = string.Join("|", filters) + "|所有文件 (*.*)|*.*";
            }
            var options = new FileSelectionOptions
            {
                Filter = filter,
                Title = "选择要导入的Excel文件"
            };
            return await ShowFileSelectionDialogAsync(options);
        }
        /// <summary>
        /// 显示文件保存对话框
        /// </summary>
        public async Task<FileSaveResult> ShowFileSaveDialogAsync(FileSaveOptions options)
        {
            try
            {
                // 使用UI线程调度器确保在正确的线程上执行
                return await UIThreadDispatcher.InvokeOnUIThreadAsync(() =>
                {
                    return ShowFileSaveDialogSync(options);
                });
            }
            catch (STAThreadException staEx)
            {
                
                var result = new FileSaveResult
                {
                    IsSuccess = false,
                    ErrorMessage = $"文件保存对话框失败: {staEx.Message}"
                };
                return result;
            }
            catch (Exception ex)
            {
                
                // 检查是否可能是STA线程问题
                if (STAThreadException.IsPotentialSTAThreadIssue(ex))
                {
                    var staEx = STAThreadException.CreateForFileDialog("文件保存对话框", ex);
                    
                    var result = new FileSaveResult
                    {
                        IsSuccess = false,
                        ErrorMessage = $"文件保存对话框失败: {staEx.Message}"
                    };
                    return result;
                }
                
                var errorResult = new FileSaveResult
                {
                    IsSuccess = false,
                    ErrorMessage = $"文件保存对话框失败: {ex.Message}"
                };
                return errorResult;
            }
        }
        /// <summary>
        /// 同步显示文件保存对话框（必须在UI线程调用）
        /// </summary>
        private FileSaveResult ShowFileSaveDialogSync(FileSaveOptions options)
        {
            var result = new FileSaveResult();
            try
            {
                using (var saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = options.Filter;
                    saveDialog.Title = options.Title;
                    saveDialog.OverwritePrompt = options.OverwritePrompt;
                    saveDialog.AddExtension = options.AddExtension;
                    if (!string.IsNullOrEmpty(options.DefaultFileName))
                    {
                        saveDialog.FileName = options.DefaultFileName;
                    }
                    if (!string.IsNullOrEmpty(options.DefaultExtension))
                    {
                        saveDialog.DefaultExt = options.DefaultExtension;
                    }
                    if (!string.IsNullOrEmpty(options.InitialDirectory))
                    {
                        saveDialog.InitialDirectory = options.InitialDirectory;
                    }
                    var dialogResult = saveDialog.ShowDialog();
                    if (dialogResult == DialogResult.OK)
                    {
                        result.IsSuccess = true;
                        result.FilePath = saveDialog.FileName;
                    }
                    else
                    {
                        result.IsCancelled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.ErrorMessage = $"文件保存对话框失败: {ex.Message}";
            }
            return result;
        }
        /// <summary>
        /// 下载文件到客户端（WinForms中实际是保存文件）
        /// </summary>
        public async Task<FileDownloadResult> DownloadFileAsync(byte[] fileContent, string fileName, string mimeType)
        {
            
            var result = new FileDownloadResult();
            try
            {
                var saveOptions = new FileSaveOptions
                {
                    DefaultFileName = fileName,
                    Title = "保存导出文件"
                };
                var saveResult = await ShowFileSaveDialogAsync(saveOptions);
                if (saveResult.IsSuccess)
                {
                    File.WriteAllBytes(saveResult.FilePath, fileContent);
                    result.IsSuccess = true;
                    // 尝试打开文件所在文件夹
                    try
                    {
                        System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{saveResult.FilePath}\"");
                    }
                    catch
                    {
                        // 忽略打开文件夹的错误
                    }
                }
                else if (saveResult.IsCancelled)
                {
                    result.ErrorMessage = "用户取消了保存操作";
                }
                else
                {
                    result.ErrorMessage = saveResult.ErrorMessage;
                }
            }
            catch (Exception ex)
            {
                result.IsSuccess = false;
                result.ErrorMessage = $"文件保存失败: {ex.Message}";
            }
            return result;
        }
        /// <summary>
        /// 显示消息
        /// </summary>
        public async Task ShowMessageAsync(string message, string title = null, MessageType messageType = MessageType.Information)
        {
            await Task.Run(() =>
            {
                try
                {
                    // 检查是否为导入完成消息
                    var isImportMessage = (messageType == MessageType.Success || messageType == MessageType.Warning) &&
                                         !string.IsNullOrEmpty(title) &&
                                         (title.Contains("导入") || title.Contains("Import"));

                    if (isImportMessage)
                    {
                        // 使用自定义的美化对话框
                        ShowImportDialog(message, title, messageType);
                    }
                    else
                    {
                        // 使用标准消息框
                        var icon = MessageBoxIcon.Information;
                        switch (messageType)
                        {
                            case MessageType.Warning:
                                icon = MessageBoxIcon.Warning;
                                break;
                            case MessageType.Error:
                                icon = MessageBoxIcon.Error;
                                break;
                            case MessageType.Success:
                                icon = MessageBoxIcon.Information;
                                break;
                        }
                        MessageBox.Show(message, title ?? "提示", MessageBoxButtons.OK, icon);
                    }
                }
                catch (Exception ex)
                {
                    // 降级到XAF消息显示
                    _application?.ShowViewStrategy?.ShowMessage($"{message}\n\n(消息显示异常: {ex.Message})");
                }
            });
        }

        /// <summary>
        /// 显示美化的导入结果对话框
        /// </summary>
        private void ShowImportDialog(string message, string title, MessageType messageType)
        {
#if NET8_0_WINDOWS || NET462_OR_GREATER
            try
            {
                using (var form = new System.Windows.Forms.Form())
                {
                    form.Text = title ?? "导入完成";
                    form.StartPosition = FormStartPosition.CenterScreen;
                    form.FormBorderStyle = FormBorderStyle.FixedDialog;
                    form.MaximizeBox = false;
                    form.MinimizeBox = false;
                    form.ClientSize = new System.Drawing.Size(680, 500);
                    form.Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Regular);

                    // 创建渐变背景的标题栏
                    var titlePanel = new System.Windows.Forms.Panel
                    {
                        Dock = DockStyle.Top,
                        Height = 70,
                        BackColor = System.Drawing.Color.FromArgb(102, 126, 234),
                        Padding = new System.Windows.Forms.Padding(20, 15, 20, 15)
                    };

                    var titleLabel = new System.Windows.Forms.Label
                    {
                        Text = title ?? "导入完成",
                        Font = new System.Drawing.Font("Microsoft YaHei UI", 14F, System.Drawing.FontStyle.Bold),
                        ForeColor = System.Drawing.Color.White,
                        Dock = DockStyle.Top,
                        Height = 30
                    };

                    var subtitleLabel = new System.Windows.Forms.Label
                    {
                        Text = "导入操作已完成",
                        Font = new System.Drawing.Font("Microsoft YaHei UI", 9F),
                        ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 230),
                        Dock = DockStyle.Top,
                        Height = 20
                    };

                    titlePanel.Controls.Add(subtitleLabel);
                    titlePanel.Controls.Add(titleLabel);

                    // 创建消息文本区域
                    var messagePanel = new System.Windows.Forms.Panel
                    {
                        Dock = DockStyle.Fill,
                        Padding = new System.Windows.Forms.Padding(20)
                    };

                    var messageTextBox = new System.Windows.Forms.TextBox
                    {
                        Text = message,
                        Multiline = true,
                        ReadOnly = true,
                        Dock = DockStyle.Fill,
                        BorderStyle = System.Windows.Forms.BorderStyle.None,
                        BackColor = System.Drawing.Color.White,
                        Font = new System.Drawing.Font("Consolas", 9F),
                        ScrollBars = System.Windows.Forms.ScrollBars.Vertical
                    };

                    messagePanel.Controls.Add(messageTextBox);

                    // 创建按钮区域
                    var buttonPanel = new System.Windows.Forms.Panel
                    {
                        Dock = DockStyle.Bottom,
                        Height = 60,
                        BackColor = System.Drawing.Color.FromArgb(250, 250, 250),
                        Padding = new System.Windows.Forms.Padding(20)
                    };

                    var copyButton = new System.Windows.Forms.Button
                    {
                        Text = "📋 复制",
                        Size = new System.Drawing.Size(100, 36),
                        UseVisualStyleBackColor = false,
                        BackColor = System.Drawing.Color.FromArgb(108, 117, 125),
                        ForeColor = System.Drawing.Color.White,
                        FlatStyle = FlatStyle.Flat,
                        Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold),
                        Cursor = System.Windows.Forms.Cursors.Hand
                    };
                    copyButton.FlatAppearance.BorderSize = 0;
                    copyButton.Click += (s, e) =>
                    {
                        System.Windows.Forms.Clipboard.SetText(message);
                        copyButton.Text = "✅ 已复制";
                        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer { Interval = 2000 };
                        timer.Tick += (ts, te) =>
                        {
                            copyButton.Text = "📋 复制";
                            timer.Stop();
                            timer.Dispose();
                        };
                        timer.Start();
                    };

                    var okButton = new System.Windows.Forms.Button
                    {
                        Text = "确定",
                        Size = new System.Drawing.Size(100, 36),
                        UseVisualStyleBackColor = false,
                        BackColor = System.Drawing.Color.FromArgb(102, 126, 234),
                        ForeColor = System.Drawing.Color.White,
                        FlatStyle = FlatStyle.Flat,
                        Font = new System.Drawing.Font("Microsoft YaHei UI", 9F, System.Drawing.FontStyle.Bold),
                        Cursor = System.Windows.Forms.Cursors.Hand,
                        DialogResult = DialogResult.OK
                    };
                    okButton.FlatAppearance.BorderSize = 0;

                    buttonPanel.Controls.Add(copyButton);
                    buttonPanel.Controls.Add(okButton);

                    // 布局按钮
                    copyButton.Location = new System.Drawing.Point(buttonPanel.ClientSize.Width - 220, 12);
                    okButton.Location = new System.Drawing.Point(buttonPanel.ClientSize.Width - 110, 12);

                    form.Controls.Add(messagePanel);
                    form.Controls.Add(buttonPanel);
                    form.Controls.Add(titlePanel);

                    form.ShowDialog();
                }
            }
            catch
            {
                // 降级到标准消息框
                MessageBox.Show(message, title ?? "导入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
#else
            MessageBox.Show(message, title ?? "导入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
#endif
        }
        /// <summary>
        /// 显示确认对话框
        /// </summary>
        public async Task<bool> ShowConfirmationAsync(string message, string title = null)
        {
            return await Task.Run(() =>
            {
                try
                {
                    var result = MessageBox.Show(message, title ?? "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    return result == DialogResult.Yes;
                }
                catch (Exception ex)
                {
                    // 降级到XAF消息显示
                    _application?.ShowViewStrategy?.ShowMessage($"{message}\n\n(确认对话框异常: {ex.Message})");
                    return false;
                }
            });
        }

        /// <summary>
        /// 显示导入选项对话框
        /// </summary>
        public async Task<ExcelImportMode?> ShowImportOptionsDialogAsync(string title = null, ExcelImportMode? defaultMode = null)
        {
            return await UIThreadDispatcher.InvokeOnUIThreadAsync(() =>
            {
                try
                {
                    // 创建自定义对话框
                    using (var form = new Form())
                    {
                        form.Text = title ?? "选择导入模式";
                        form.FormBorderStyle = FormBorderStyle.FixedDialog;
                        form.MaximizeBox = false;
                        form.MinimizeBox = false;
                        form.StartPosition = FormStartPosition.CenterScreen;
                        form.ClientSize = new System.Drawing.Size(450, 320);
                        form.BackColor = System.Drawing.Color.White;

                        // 标题标签
                        var titleLabel = new Label
                        {
                            Text = "请选择数据导入模式:",
                            Font = new System.Drawing.Font("Microsoft YaHei UI", 10F, System.Drawing.FontStyle.Bold),
                            Location = new System.Drawing.Point(20, 20),
                            Size = new System.Drawing.Size(400, 30),
                            ForeColor = System.Drawing.Color.FromArgb(51, 51, 51)
                        };
                        form.Controls.Add(titleLabel);

                        // 创建选项单选按钮
                        var radioButton1 = new RadioButton
                        {
                            Text = "仅新增 (Insert)",
                            Location = new System.Drawing.Point(20, 60),
                            Size = new System.Drawing.Size(400, 24),
                            Tag = ExcelImportMode.CreateOnly
                        };
                        var descLabel1 = new Label
                        {
                            Text = "只创建新记录,如果记录已存在则跳过",
                            Location = new System.Drawing.Point(40, 84),
                            Size = new System.Drawing.Size(380, 20),
                            ForeColor = System.Drawing.Color.FromArgb(102, 102, 102),
                            Font = new System.Drawing.Font("Microsoft YaHei UI", 8F)
                        };

                        var radioButton2 = new RadioButton
                        {
                            Text = "仅更新 (Update)",
                            Location = new System.Drawing.Point(20, 110),
                            Size = new System.Drawing.Size(400, 24),
                            Tag = ExcelImportMode.UpdateOnly
                        };
                        var descLabel2 = new Label
                        {
                            Text = "只更新现有记录,如果记录不存在则跳过",
                            Location = new System.Drawing.Point(40, 134),
                            Size = new System.Drawing.Size(380, 20),
                            ForeColor = System.Drawing.Color.FromArgb(102, 102, 102),
                            Font = new System.Drawing.Font("Microsoft YaHei UI", 8F)
                        };

                        var radioButton3 = new RadioButton
                        {
                            Text = "新增或更新 (InsertOrUpdate)",
                            Location = new System.Drawing.Point(20, 160),
                            Size = new System.Drawing.Size(400, 24),
                            Tag = ExcelImportMode.CreateOrUpdate
                        };
                        var descLabel3 = new Label
                        {
                            Text = "存在则更新,不存在则新增 (推荐)",
                            Location = new System.Drawing.Point(40, 184),
                            Size = new System.Drawing.Size(380, 20),
                            ForeColor = System.Drawing.Color.FromArgb(102, 102, 102),
                            Font = new System.Drawing.Font("Microsoft YaHei UI", 8F)
                        };

                        var radioButton4 = new RadioButton
                        {
                            Text = "替换全部 (ReplaceAll)",
                            Location = new System.Drawing.Point(20, 210),
                            Size = new System.Drawing.Size(400, 24),
                            Tag = ExcelImportMode.ReplaceAll
                        };
                        var descLabel4 = new Label
                        {
                            Text = "删除所有现有记录后重新导入",
                            Location = new System.Drawing.Point(40, 234),
                            Size = new System.Drawing.Size(380, 20),
                            ForeColor = System.Drawing.Color.FromArgb(102, 102, 102),
                            Font = new System.Drawing.Font("Microsoft YaHei UI", 8F)
                        };

                        // 设置默认选中
                        var defaultRadio = radioButton3;
                        if (defaultMode.HasValue)
                        {
                            switch (defaultMode.Value)
                            {
                                case ExcelImportMode.CreateOnly:
                                    defaultRadio = radioButton1;
                                    break;
                                case ExcelImportMode.UpdateOnly:
                                    defaultRadio = radioButton2;
                                    break;
                                case ExcelImportMode.ReplaceAll:
                                    defaultRadio = radioButton4;
                                    break;
                                default:
                                    defaultRadio = radioButton3;
                                    break;
                            }
                        }
                        defaultRadio.Checked = true;

                        // 添加控件到表单
                        form.Controls.Add(radioButton1);
                        form.Controls.Add(descLabel1);
                        form.Controls.Add(radioButton2);
                        form.Controls.Add(descLabel2);
                        form.Controls.Add(radioButton3);
                        form.Controls.Add(descLabel3);
                        form.Controls.Add(radioButton4);
                        form.Controls.Add(descLabel4);

                        // 按钮
                        var okButton = new Button
                        {
                            Text = "确定",
                            DialogResult = DialogResult.OK,
                            Location = new System.Drawing.Point(260, 265),
                            Size = new System.Drawing.Size(80, 30),
                            BackColor = System.Drawing.Color.FromArgb(0, 120, 212),
                            ForeColor = System.Drawing.Color.White,
                            FlatStyle = FlatStyle.Flat,
                            Font = new System.Drawing.Font("Microsoft YaHei UI", 9F)
                        };
                        okButton.FlatAppearance.BorderSize = 0;

                        var cancelButton = new Button
                        {
                            Text = "取消",
                            DialogResult = DialogResult.Cancel,
                            Location = new System.Drawing.Point(350, 265),
                            Size = new System.Drawing.Size(80, 30),
                            BackColor = System.Drawing.Color.FromArgb(108, 117, 125),
                            ForeColor = System.Drawing.Color.White,
                            FlatStyle = FlatStyle.Flat,
                            Font = new System.Drawing.Font("Microsoft YaHei UI", 9F)
                        };
                        cancelButton.FlatAppearance.BorderSize = 0;

                        form.Controls.Add(okButton);
                        form.Controls.Add(cancelButton);
                        form.AcceptButton = okButton;
                        form.CancelButton = cancelButton;

                        // 显示对话框
                        var result = form.ShowDialog();

                        if (result == DialogResult.OK)
                        {
                            if (radioButton1.Checked) return (ExcelImportMode?)ExcelImportMode.CreateOnly;
                            if (radioButton2.Checked) return (ExcelImportMode?)ExcelImportMode.UpdateOnly;
                            if (radioButton3.Checked) return (ExcelImportMode?)ExcelImportMode.CreateOrUpdate;
                            if (radioButton4.Checked) return (ExcelImportMode?)ExcelImportMode.ReplaceAll;
                        }

                        return (ExcelImportMode?)null; // 用户取消
                    }
                }
                catch (Exception ex)
                {
                    // 降级: 返回默认模式
                    _application?.ShowViewStrategy?.ShowMessage($"导入选项对话框异常: {ex.Message}\n\n将使用默认模式");
                    return defaultMode ?? ExcelImportMode.CreateOrUpdate;
                }
            });
        }

        /// <summary>
        /// 导航到指定页面（WinForms中不适用）
        /// </summary>
        public async Task<NavigationResult> NavigateToAsync(string url, bool newWindow = false)
        {
            return await Task.Run(() =>
            {
                var result = new NavigationResult();
                try
                {
                    if (newWindow)
                    {
                        // 在外部浏览器中打开
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = url,
                            UseShellExecute = true
                        });
                        result.IsSuccess = true;
                    }
                    else
                    {
                        result.IsSuccess = false;
                        result.ErrorMessage = "WinForms应用程序不支持内部页面导航";
                    }
                }
                catch (Exception ex)
                {
                    result.IsSuccess = false;
                    result.ErrorMessage = $"导航失败: {ex.Message}";
                }
                return result;
            });
        }
        /// <summary>
        /// 根据文件扩展名获取MIME类型
        /// </summary>
        private string GetMimeType(string extension)
        {
            return extension?.ToLowerInvariant() switch
            {
                ".csv" => "text/csv",
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ".xls" => "application/vnd.ms-excel",
                _ => "application/octet-stream"
            };
        }
        /// <summary>
        /// 获取平台文件服务的诊断信息
        /// </summary>
        /// <returns>诊断信息字符串</returns>
        public string GetDiagnosticInfo()
        {
            var info = new System.Text.StringBuilder();
            info.AppendLine("=== WinForms平台文件服务诊断信息 ===");
            
            try
            {
                info.AppendLine($"Application: {(_application != null ? "可用" : "null")}");
                info.AppendLine(ThreadStateChecker.GetThreadDiagnosticInfo());
                info.AppendLine(UIThreadDispatcher.GetDiagnosticInfo());
                
                var envCheck = ThreadStateChecker.CheckFileDialogEnvironment();
                info.AppendLine(envCheck.GetReport());
            }
            catch (Exception ex)
            {
                info.AppendLine($"诊断信息收集失败: {ex.Message}");
            }
            
            info.AppendLine("================================");
            return info.ToString();
        }
    }
}

#endif // NET8_0_WINDOWS || NET462_OR_GREATER
