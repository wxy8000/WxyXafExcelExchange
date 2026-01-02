using System;
using System.Threading;
using System.Threading.Tasks;

#if NET8_0_WINDOWS || NET462_OR_GREATER
using System.Windows.Forms;

namespace Wxy.Xaf.ExcelExchange.Threading
{
    /// <summary>
    /// UI线程调度器，用于确保操作在正确的线程上执行
    /// 仅在Windows平台可用
    /// </summary>
    public static class UIThreadDispatcher
    {
        private static Control _uiControl;
        private static readonly object _lockObject = new object();

        /// <summary>
        /// 在UI线程上异步执行带返回值的操作
        /// </summary>
        /// <typeparam name="T">返回值类型</typeparam>
        /// <param name="action">要执行的操作</param>
        /// <returns>操作结果</returns>
        public static async Task<T> InvokeOnUIThreadAsync<T>(Func<T> action)
        {
            if (action == null)
                throw new ArgumentNullException(nameof(action));

            var context = FileDialogExecutionContext.Create($"InvokeOnUIThreadAsync<{typeof(T).Name}>");
            
            try
            {
                // 如果已经在STA线程上，直接执行
                if (ThreadStateChecker.IsSTAThread())
                {
                    context.MarkExecutionStart();
                    var result = action();
                    context.MarkExecutionComplete(true);
                    
                                        return result;
                }

                // 获取UI控件用于调度
                var uiControl = GetOrCreateUIControl();
                if (uiControl == null)
                {
                    throw STAThreadException.CreateForFileDialog("UI线程调度");
                }

                // 使用TaskCompletionSource来处理异步调用
                var tcs = new TaskCompletionSource<T>();

                uiControl.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        context.MarkExecutionStart();
                        var result = action();
                        context.MarkExecutionComplete(true);
                        tcs.SetResult(result);
                    }
                    catch (Exception ex)
                    {
                        context.MarkExecutionComplete(false, ex.Message);
                        tcs.SetException(ex);
                    }
                }));

                var finalResult = await tcs.Task;
                                return finalResult;
            }
            catch (Exception ex)
            {
                context.MarkExecutionComplete(false, ex.Message);
                                
                // 如果是STA相关异常，直接抛出
                if (ex is STAThreadException)
                    throw;
                
                // 检查是否可能是STA线程问题
                if (STAThreadException.IsPotentialSTAThreadIssue(ex))
                {
                    throw STAThreadException.CreateForFileDialog("UI线程调度", ex);
                }
                
                throw;
            }
        }

        /// <summary>
        /// 在UI线程上异步执行无返回值的操作
        /// </summary>
        /// <param name="action">要执行的操作</param>
        /// <returns>异步任务</returns>
        public static async Task InvokeOnUIThreadAsync(Action action)
        {
            if (action == null)
                throw new ArgumentNullException(nameof(action));

            await InvokeOnUIThreadAsync(() =>
            {
                action();
                return true; // 返回一个虚拟值
            });
        }

        /// <summary>
        /// 尝试获取UI线程引用
        /// </summary>
        /// <param name="uiThread">UI线程引用</param>
        /// <returns>如果成功获取返回true</returns>
        public static bool TryGetUIThread(out Thread uiThread)
        {
            uiThread = null;
            
            try
            {
                var uiControl = GetOrCreateUIControl();
                if (uiControl != null && uiControl.IsHandleCreated)
                {
                    // 通过Invoke获取UI线程
                    Thread capturedThread = null;
                    uiControl.Invoke(new Action(() =>
                    {
                        capturedThread = Thread.CurrentThread;
                    }));
                    
                    uiThread = capturedThread;
                    return uiThread != null;
                }
            }
            catch
            {
                            }

            return false;
        }

        /// <summary>
        /// 获取或创建UI控件用于线程调度
        /// </summary>
        /// <returns>UI控件实例</returns>
        private static Control GetOrCreateUIControl()
        {
            if (_uiControl != null && !_uiControl.IsDisposed && _uiControl.IsHandleCreated)
            {
                return _uiControl;
            }

            lock (_lockObject)
            {
                // 双重检查锁定
                if (_uiControl != null && !_uiControl.IsDisposed && _uiControl.IsHandleCreated)
                {
                    return _uiControl;
                }

                try
                {
                    // 首先尝试使用现有的窗体
                    if (Application.OpenForms.Count > 0)
                    {
                        var mainForm = Application.OpenForms[0];
                        if (mainForm != null && !mainForm.IsDisposed && mainForm.IsHandleCreated)
                        {
                            _uiControl = mainForm;
                                                        return _uiControl;
                        }
                    }

                    // 如果没有现有窗体，创建一个隐藏的控件
                    _uiControl = new Control();
                    _uiControl.CreateControl(); // 强制创建句柄
                    
                                        return _uiControl;
                }
                catch
                {
                                        _uiControl = null;
                    return null;
                }
            }
        }

        /// <summary>
        /// 检查当前是否可以进行UI线程调度
        /// </summary>
        /// <returns>如果可以调度返回true</returns>
        public static bool CanDispatchToUIThread()
        {
            try
            {
                var uiControl = GetOrCreateUIControl();
                return uiControl != null && !uiControl.IsDisposed && uiControl.IsHandleCreated;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 获取UI线程调度器的诊断信息
        /// </summary>
        /// <returns>诊断信息字符串</returns>
        public static string GetDiagnosticInfo()
        {
            var info = new System.Text.StringBuilder();
            info.AppendLine("=== UI线程调度器诊断信息 ===");
            
            try
            {
                info.AppendLine($"当前线程: {ThreadDiagnosticInfo.CreateForCurrentThread()}");
                info.AppendLine($"可以调度到UI线程: {CanDispatchToUIThread()}");
                
                var uiControl = _uiControl;
                if (uiControl != null)
                {
                    info.AppendLine($"UI控件类型: {uiControl.GetType().Name}");
                    info.AppendLine($"UI控件已释放: {uiControl.IsDisposed}");
                    info.AppendLine($"UI控件句柄已创建: {uiControl.IsHandleCreated}");
                    
                    if (!uiControl.IsDisposed)
                    {
                        try
                        {
                            info.AppendLine($"需要调用: {uiControl.InvokeRequired}");
                        }
                        catch (Exception ex)
                        {
                            info.AppendLine($"检查InvokeRequired失败: {ex.Message}");
                        }
                    }
                }
                else
                {
                    info.AppendLine("UI控件: null");
                }
                
                info.AppendLine($"WinForms窗体数量: {Application.OpenForms.Count}");
                
                if (TryGetUIThread(out Thread uiThread))
                {
                    info.AppendLine($"UI线程ID: {uiThread.ManagedThreadId}");
                    info.AppendLine($"UI线程状态: {uiThread.GetApartmentState()}");
                }
                else
                {
                    info.AppendLine("无法获取UI线程引用");
                }
            }
            catch (Exception ex)
            {
                info.AppendLine($"诊断信息收集失败: {ex.Message}");
            }
            
            info.AppendLine("===========================");
            return info.ToString();
        }

        /// <summary>
        /// 清理资源
        /// </summary>
        public static void Cleanup()
        {
            lock (_lockObject)
            {
                if (_uiControl != null && !_uiControl.IsDisposed)
                {
                    try
                    {
                        _uiControl.Dispose();
                    }
                    catch
                    {
                                            }
                }
                _uiControl = null;
            }
        }
    }
}

#endif // NET8_0_WINDOWS || NET462_OR_GREATER
