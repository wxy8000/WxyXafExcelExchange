using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Actions;
using DevExpress.Xpo;
using Wxy.Xaf.ExcelExchange;
using Wxy.Xaf.ExcelExchange.Services;
using Wxy.Xaf.ExcelExchange.Configuration;
using System.Diagnostics;

namespace Wxy.Xaf.ExcelExchange.Controllers
{
    /// <summary>
    /// Excel导入导出控制器基类（平台无关的核心逻辑）
    /// 支持新的双层特性系统
    /// </summary>
    public abstract class ExcelImportExportControllerBase : ViewController<ListView>
    {
        #region 字段和属性

        protected SimpleAction _exportAction;
        protected SimpleAction _importAction;
        protected IExcelImportExportService _excelService;
        protected IPlatformFileService _platformFileService;
        protected ConfigurationManager _configurationManager;

        /// <summary>
        /// Excel服务
        /// </summary>
        protected IExcelImportExportService ExcelService
        {
            get
            {
                if (_excelService == null)
                {
                    _excelService = GetExcelService();
                }
                return _excelService;
            }
        }

        /// <summary>
        /// 平台文件服务
        /// </summary>
        protected IPlatformFileService PlatformFileService
        {
            get
            {
                if (_platformFileService == null)
                {
                    _platformFileService = GetPlatformFileService();
                }
                return _platformFileService;
            }
        }

        /// <summary>
        /// 配置管理器
        /// </summary>
        protected ConfigurationManager ConfigurationManager
        {
            get
            {
                if (_configurationManager == null)
                {
                    _configurationManager = new ConfigurationManager();
                }
                return _configurationManager;
            }
        }

        #endregion

        #region 构造函数和初始化

        protected ExcelImportExportControllerBase()
        {
            InitializeActions();
        }

        /// <summary>
        /// 初始化操作
        /// </summary>
        protected virtual void InitializeActions()
        {
            // 获取平台特定的操作ID前缀
            var actionIdPrefix = GetActionIdPrefix();

            // 初始化导出操作 - 使用平台特定的ID避免冲突
            // 放在查看工具栏（View）而不是导出菜单（Export）
            _exportAction = new SimpleAction(this, $"{actionIdPrefix}ExportToExcelV2", DevExpress.Persistent.Base.PredefinedCategory.View)
            {
                Caption = "导出到Excel",
                ImageName = "Export", // 使用DevExpress标准图标
                ToolTip = "导出当前列表数据到Excel文件"
            };
            _exportAction.Active["HasExcelAttribute"] = false;
            _exportAction.Execute += ExportAction_Execute;

            // 初始化导入操作 - 使用平台特定的ID避免冲突
            // 放在查看工具栏（View）而不是导出菜单（Export）
            _importAction = new SimpleAction(this, $"{actionIdPrefix}ImportFromExcelV2", DevExpress.Persistent.Base.PredefinedCategory.View)
            {
                Caption = "从Excel导入",
                ImageName = "Import", // 使用DevExpress标准图标
                ToolTip = "从Excel文件导入数据到当前列表"
            };
            _importAction.Active["HasExcelAttribute"] = false;
            _importAction.Execute += ImportAction_Execute;
        }

        /// <summary>
        /// 获取操作ID前缀（由子类实现以确保唯一性）
        /// </summary>
        protected abstract string GetActionIdPrefix();

        protected override void OnViewChanged()
        {
            base.OnViewChanged();
            UpdateActionAvailability();
        }

        /// <summary>
        /// 更新操作可用性
        /// </summary>
        protected virtual void UpdateActionAvailability()
        {
            try
            {
                if (View?.ObjectTypeInfo?.Type == null)
                {
                    _exportAction.Active.SetItemValue("HasExcelAttribute", false);
                    _importAction.Active.SetItemValue("HasExcelAttribute", false);
                    return;
                }

                var objectType = View.ObjectTypeInfo.Type;

                // 使用新的配置管理器检查Excel支持
                bool hasExcelSupport = ConfigurationManager.IsExcelSupported(objectType);

                // 检查是否继承自XPObject或具有XPO特征（如Session属性）
                bool isXPOObject = typeof(XPObject).IsAssignableFrom(objectType) ||
                                  HasXPOCharacteristics(objectType);

                // 设置操作可用性
                bool actionsEnabled = hasExcelSupport && isXPOObject;

                _exportAction.Active.SetItemValue("HasExcelAttribute", actionsEnabled);
                _importAction.Active.SetItemValue("HasExcelAttribute", actionsEnabled);

                // 如果支持Excel，获取配置并自定义操作标题
                if (hasExcelSupport)
                {
                    try
                    {
                        var config = ConfigurationManager.GetConfiguration(objectType, View.ObjectTypeInfo);
                        if (config.ClassConfiguration != null)
                        {
                            _exportAction.Caption = config.ClassConfiguration.ExportCaption ?? "导出到Excel";
                            _importAction.Caption = config.ClassConfiguration.ImportCaption ?? "从Excel导入";

                            // 设置操作描述
                            if (!string.IsNullOrEmpty(config.ClassConfiguration.Description))
                            {
                                _exportAction.ToolTip = config.ClassConfiguration.Description;
                                _importAction.ToolTip = config.ClassConfiguration.Description;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogError($"获取Excel配置失败: {ex.Message}", ex);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"更新操作可用性失败: {ex.Message}", ex);
            }
        }

        #endregion

        #region 导出逻辑

        /// <summary>
        /// 导出操作执行
        /// </summary>
        protected virtual async void ExportAction_Execute(object sender, SimpleActionExecuteEventArgs e)
        {
            try
            {
                LogDebug("开始导出操作");

                if (!ValidateViewForExport())
                {
                    return;
                }

                var objectType = View.ObjectTypeInfo.Type;
                var data = GetExportData();

                if (!data.Any())
                {
                    await PlatformFileService.ShowMessageAsync($"暂无{objectType.Name}数据可导出", "提示", MessageType.Warning);
                    return;
                }

                LogDebug($"准备导出 {data.Count()} 条 {objectType.Name} 数据");

                // 获取导出选项
                var exportOptions = GetExportOptions(objectType);

                // 执行导出
                var exportResult = await ExcelService.ExportDataAsync(data, objectType, exportOptions);

                if (exportResult.IsSuccess)
                {
                    // 下载文件
                    var downloadResult = await PlatformFileService.DownloadFileAsync(
                        exportResult.FileContent,
                        exportResult.SuggestedFileName,
                        exportResult.MimeType);

                    if (downloadResult.IsSuccess)
                    {
                        //await PlatformFileService.ShowMessageAsync(
                        //    $"导出成功！共导出 {exportResult.RecordCount} 条记录",
                        //    "导出完成",
                        //    MessageType.Success);
                        LogDebug($"导出成功: {exportResult.RecordCount} 条记录");
                    }
                    else
                    {
                        await PlatformFileService.ShowMessageAsync(
                            $"文件下载失败: {downloadResult.ErrorMessage}",
                            "导出失败",
                            MessageType.Error);
                        LogError($"文件下载失败: {downloadResult.ErrorMessage}");
                    }
                }
                else
                {
                    await PlatformFileService.ShowMessageAsync(
                        $"导出失败: {exportResult.ErrorMessage}",
                        "导出失败",
                        MessageType.Error);
                    LogError($"导出失败: {exportResult.ErrorMessage}");
                }
            }
            catch (Exception ex)
            {
                LogError($"导出操作异常: {ex.Message}", ex);
                await PlatformFileService.ShowMessageAsync(
                    $"导出失败: {ex.Message}",
                    "导出异常",
                    MessageType.Error);
            }
        }

        /// <summary>
        /// 验证视图是否可以导出
        /// </summary>
        protected virtual bool ValidateViewForExport()
        {
            if (View?.ObjectSpace == null || View.ObjectTypeInfo?.Type == null)
            {
                LogError("视图或对象空间为空，无法导出");
                return false;
            }

            return true;
        }

        /// <summary>
        /// 获取要导出的数据
        /// </summary>
        protected virtual IEnumerable<object> GetExportData()
        {
            try
            {
                var objectType = View.ObjectTypeInfo.Type;
                var objectSpace = View.ObjectSpace;

                // 使用反射获取对象列表
                var getObjectsMethod = objectSpace.GetType().GetMethod("GetObjects", Type.EmptyTypes);
                var genericGetObjectsMethod = getObjectsMethod.MakeGenericMethod(objectType);
                var dataCollection = genericGetObjectsMethod.Invoke(objectSpace, null);

                // 转换为可枚举对象
                if (dataCollection is System.Collections.IEnumerable enumerable)
                {
                    return enumerable.Cast<object>().ToList();
                }

                return new List<object>();
            }
            catch (Exception ex)
            {
                LogError($"获取导出数据失败: {ex.Message}", ex);
                return new List<object>();
            }
        }

        /// <summary>
        /// 获取导出选项
        /// </summary>
        protected virtual ExcelExportOptions GetExportOptions(Type objectType)
        {
            try
            {
                var config = ConfigurationManager.GetConfiguration(objectType, View?.ObjectTypeInfo);
                var classConfig = config.ClassConfiguration;
                
                return new ExcelExportOptions
                {
                    Format = ExcelFormat.Xlsx, // 默认使用 XLSX 格式
                    FileName = classConfig?.GetEffectiveDefaultFileName(objectType) ?? objectType.Name,
                    IncludeHeaders = classConfig?.ExportIncludeHeader ?? true,
                    Encoding = Encoding.UTF8, // 统一使用UTF-8编码
                    AddBom = true // 添加BOM确保兼容性
                };
            }
            catch (Exception ex)
            {
                LogError($"获取导出选项失败: {ex.Message}", ex);

                // 返回默认选项
                return new ExcelExportOptions
                {
                    Format = ExcelFormat.Xlsx,
                    FileName = objectType.Name,
                    IncludeHeaders = true,
                    Encoding = Encoding.UTF8,
                    AddBom = true
                };
            }
        }

        #endregion

        #region 导入逻辑

        /// <summary>
        /// 导入操作执行
        /// </summary>
        protected virtual async void ImportAction_Execute(object sender, SimpleActionExecuteEventArgs e)
        {
            try
            {
                LogDebug("开始导入操作");

                if (!ValidateViewForImport())
                {
                    return;
                }

                var objectType = View.ObjectTypeInfo.Type;
                LogDebug($"导入对象类型: {objectType.FullName}");

                // 执行平台特定的导入流程
                await ExecutePlatformSpecificImport(objectType);
            }
            catch (Exception ex)
            {
                LogError($"导入操作异常: {ex.Message}", ex);
                await PlatformFileService.ShowMessageAsync(
                    $"导入失败: {ex.Message}",
                    "导入异常",
                    MessageType.Error);
            }
        }

        /// <summary>
        /// 验证视图是否可以导入
        /// </summary>
        protected virtual bool ValidateViewForImport()
        {
            if (View?.ObjectSpace == null || View.ObjectTypeInfo?.Type == null)
            {
                LogError("视图或对象空间为空，无法导入");
                return false;
            }

            return true;
        }

        /// <summary>
        /// 执行平台特定的导入流程（由子类实现）
        /// </summary>
        protected abstract Task ExecutePlatformSpecificImport(Type objectType);

        /// <summary>
        /// 处理文件导入（平台无关的核心逻辑）
        /// </summary>
        protected virtual async Task<ExcelImportResult> ProcessFileImport(byte[] fileContent, string fileName, Type objectType)
        {
            return await ProcessFileImport(fileContent, fileName, objectType, null);
        }

        /// <summary>
        /// 处理文件导入（平台无关的核心逻辑，支持指定导入模式）
        /// </summary>
        protected virtual async Task<ExcelImportResult> ProcessFileImport(byte[] fileContent, string fileName, Type objectType, ExcelImportMode? importMode)
        {
            try
            {
                LogDebug($"开始处理文件导入: {fileName}, 大小: {fileContent?.Length ?? 0} 字节");
                if (importMode.HasValue)
                {
                    LogDebug($"指定的导入模式: {importMode.Value}");
                }

                // 验证文件
                var validation = ExcelService.ValidateFile(fileName, fileContent);
                if (!validation.IsValid)
                {
                    var result = new ExcelImportResult
                    {
                        IsSuccess = false,
                        ErrorMessage = validation.ErrorMessage
                    };
                    result.Errors.Add(new ExcelImportError
                    {
                        RowNumber = 0,
                        FieldName = "文件验证",
                        ErrorMessage = validation.ErrorMessage,
                        ErrorType = ExcelImportErrorType.SystemError
                    });
                    return result;
                }

                // 获取导入选项
                var importOptions = GetImportOptions(objectType);

                // **如果指定了导入模式,覆盖默认设置**
                if (importMode.HasValue)
                {
                    importOptions.Mode = importMode.Value;
                    importOptions.IsUserSpecifiedMode = true; // **标记为用户显式指定**
                    LogDebug($"使用用户指定的导入模式: {importMode.Value}");
                }

                // 执行导入 - 传递文件名以正确识别文件格式
                var importResult = await ExcelService.ImportDataAsync(fileContent, objectType, View.ObjectSpace, importOptions, fileName);

                // 如果导入成功，提交更改
                if (importResult.IsSuccess && importResult.SuccessCount > 0)
                {
                    try
                    {
                        LogDebug($"准备提交 {importResult.SuccessCount} 条记录的更改");

                        // **Blazor 特殊处理**: 检测是否为 Blazor 平台
                        bool isBlazor = this.GetType().Name.Contains("Blazor");

                        if (isBlazor)
                        {
                            // **Blazor**: 直接提交,不使用 Task.Run
                            // 在 Blazor 中,Task.Run 会导致跨线程问题,因为后续的 UI 刷新会在不同的线程上访问 UnitOfWork
                            // 正确的做法是让所有操作都在同一个线程上下文中执行
                            LogDebug("[Blazor] 直接提交更改（不使用 Task.Run）");
                            View.ObjectSpace.CommitChanges();
                        }
                        else
                        {
                            // WinForms: 直接提交
                            View.ObjectSpace.CommitChanges();
                        }

                        // **Blazor 特殊处理**: 跳过 Refresh(),避免跨线程验证错误
                        // Blazor 中的 Refresh() 会触发关联对象加载,导致跨线程问题
                        // CommitChanges() 已经保存了数据,Refresh() 是可选的
                        // Blazor Grid 会在需要时自动重新加载数据
                        if (!isBlazor)
                        {
                            View.ObjectSpace.Refresh();
                        }
                        else
                        {
                            LogDebug("[Blazor] 跳过 ObjectSpace.Refresh() 以避免跨线程问题");
                        }

                        LogDebug($"导入成功: {importResult.SuccessCount} 条记录已保存");
                    }
                    catch (Exception commitEx)
                    {
                        LogError($"提交更改失败: {commitEx.Message}", commitEx);

                        // 修改导入结果为失败
                        importResult.IsSuccess = false;
                        importResult.ErrorMessage = $"数据验证失败: {commitEx.Message}";
                        importResult.Errors.Add(new ExcelImportError
                        {
                            RowNumber = 0,
                            FieldName = "数据提交",
                            ErrorMessage = commitEx.Message,
                            ErrorType = ExcelImportErrorType.SystemError
                        });

                        // 回滚更改
                        try
                        {
                            View.ObjectSpace.Rollback();
                        }
                        catch (Exception rollbackEx)
                        {
                            LogError($"回滚失败: {rollbackEx.Message}", rollbackEx);
                        }
                    }
                }

                return importResult;
            }
            catch (Exception ex)
            {
                LogError($"文件导入处理失败: {ex.Message}", ex);

                // 回滚更改
                try
                {
                    View.ObjectSpace.Rollback();
                }
                catch (Exception rollbackEx)
                {
                    LogError($"回滚失败: {rollbackEx.Message}", rollbackEx);
                }

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
        /// 获取导入选项
        /// </summary>
        protected virtual ExcelImportOptions GetImportOptions(Type objectType)
        {
            try
            {
                var config = ConfigurationManager.GetConfiguration(objectType, View?.ObjectTypeInfo);
                var classConfig = config.ClassConfiguration;
                
                return new ExcelImportOptions
                {
                    Mode = ExcelImportMode.CreateOrUpdate,
                    SkipDuplicates = true,
                    HasHeaderRow = true,
                    MaxErrors = 100,
                    BatchSize = 1000
                };
            }
            catch (Exception ex)
            {
                LogError($"获取导入选项失败: {ex.Message}", ex);
                
                // 返回默认选项
                return new ExcelImportOptions
                {
                    Mode = ExcelImportMode.CreateOrUpdate,
                    SkipDuplicates = true,
                    HasHeaderRow = true,
                    MaxErrors = 100,
                    BatchSize = 1000
                };
            }
        }

        /// <summary>
        /// 显示导入结果
        /// </summary>
        protected virtual async Task ShowImportResult(ExcelImportResult result)
        {
            if (result.IsSuccess)
            {
                var message = $"导入完成！\n\n" +
                             $"总记录数: {result.TotalRecords}\n" +
                             $"成功: {result.SuccessCount}\n" +
                             $"失败: {result.FailureCount}";

                if (result.Warnings.Count > 0)
                {
                    message += $"\n警告: {result.Warnings.Count}";
                }

                await PlatformFileService.ShowMessageAsync(message, "导入完成", MessageType.Success);
            }
            else
            {
                var message = $"导入失败！\n\n错误信息: {result.ErrorMessage}";
                
                if (result.Errors.Count > 0)
                {
                    message += $"\n\n详细错误 (前5个):\n";
                    var topErrors = result.Errors.Take(5);
                    foreach (var error in topErrors)
                    {
                        message += $"行{error.RowNumber}: {error.FieldName} - {error.ErrorMessage}\n";
                    }
                    
                    if (result.Errors.Count > 5)
                    {
                        message += $"... 还有 {result.Errors.Count - 5} 个错误";
                    }
                }

                await PlatformFileService.ShowMessageAsync(message, "导入失败", MessageType.Error);
            }
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 检查类型是否具有XPO特征（如Session属性）
        /// </summary>
        protected virtual bool HasXPOCharacteristics(Type objectType)
        {
            try
            {
                // 检查是否有Session属性（XPO对象的标志）
                var sessionProperty = objectType.GetProperty("Session");
                if (sessionProperty != null && sessionProperty.PropertyType.Name == "Session")
                {
                    return true;
                }

                // 检查是否有Oid属性（XAF BaseObject的标志）
                var oidProperty = objectType.GetProperty("Oid");
                if (oidProperty != null && oidProperty.PropertyType == typeof(Guid))
                {
                    return true;
                }

                // 检查基类名称是否包含XPO相关的名称
                var baseType = objectType.BaseType;
                while (baseType != null && baseType != typeof(object))
                {
                    if (baseType.Name == "BaseObject" || baseType.Name == "XPObject" || 
                        baseType.Name.Contains("XPO") || baseType.Name.Contains("BaseObject"))
                    {
                        return true;
                    }
                    baseType = baseType.BaseType;
                }

                return false;
            }
            catch (Exception ex)
            {
                LogError($"检查XPO特征失败: {ex.Message}", ex);
                return false;
            }
        }

        /// <summary>
        /// 获取Excel导入导出特性
        /// </summary>
        protected virtual ExcelImportExportAttribute GetExcelImportExportAttribute(Type objectType)
        {
            try
            {
                // 优先从XAF的TypeInfo中获取
                var attrFromTypeInfo = View?.ObjectTypeInfo?.FindAttribute<ExcelImportExportAttribute>();
                if (attrFromTypeInfo != null)
                {
                    return attrFromTypeInfo;
                }

                // 从类型直接获取
                var attrDirect = objectType.GetCustomAttributes(typeof(ExcelImportExportAttribute), true)
                    .FirstOrDefault() as ExcelImportExportAttribute;
                
                return attrDirect;
            }
            catch (Exception ex)
            {
                LogError($"获取Excel特性失败: {ex.Message}", ex);
                return null;
            }
        }

        /// <summary>
        /// 获取Excel服务（由子类实现或通过依赖注入）
        /// </summary>
        protected abstract IExcelImportExportService GetExcelService();

        /// <summary>
        /// 获取平台文件服务（由子类实现）
        /// </summary>
        protected abstract IPlatformFileService GetPlatformFileService();

        /// <summary>
        /// 记录调试信息
        /// </summary>
        protected virtual void LogDebug(string message)
        {
                    }

        /// <summary>
        /// 记录错误信息
        /// </summary>
        protected virtual void LogError(string message, Exception exception = null)
        {
                    }

        #endregion
    }
}