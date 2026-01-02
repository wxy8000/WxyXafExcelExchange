using System;
using System.Threading;
using System.Windows.Forms;

namespace Wxy.Xaf.ExcelExchange.Threading
{
    /// <summary>
    /// 线程状态检查和诊断工具
    /// </summary>
    public static class ThreadStateChecker
    {
        /// <summary>
        /// 检查当前线程是否为STA模式
        /// </summary>
        /// <returns>如果当前线程是STA模式返回true，否则返回false</returns>
        public static bool IsSTAThread()
        {
            return Thread.CurrentThread.GetApartmentState() == ApartmentState.STA;
        }

        /// <summary>
        /// 检查UI线程是否可用
        /// </summary>
        /// <returns>如果UI线程可用返回true，否则返回false</returns>
        public static bool IsUIThreadAvailable()
        {
            try
            {
                // 检查是否有WinForms应用程序实例
                if (Application.OpenForms.Count > 0)
                {
                    return true;
                }

                // 检查是否可以创建Control（这会检查UI线程）
                using (var control = new Control())
                {
                    return control.Handle != IntPtr.Zero;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 确保当前线程为STA模式（如果可能的话）
        /// </summary>
        /// <returns>如果成功设置为STA模式或已经是STA模式返回true，否则返回false</returns>
        public static bool EnsureSTAThread()
        {
            try
            {
                // 如果已经是STA模式，直接返回true
                if (IsSTAThread())
                {
                    return true;
                }

                // 尝试设置当前线程为STA模式
                Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
                
                // 验证设置是否成功
                return IsSTAThread();
            }
            catch
            {
                                return false;
            }
        }

        /// <summary>
        /// 获取当前线程的详细诊断信息
        /// </summary>
        /// <returns>包含线程状态信息的字符串</returns>
        public static string GetThreadDiagnosticInfo()
        {
            var currentThread = Thread.CurrentThread;
            var info = new System.Text.StringBuilder();
            
            try
            {
                info.AppendLine("=== 线程诊断信息 ===");
                info.AppendLine($"线程ID: {currentThread.ManagedThreadId}");
                info.AppendLine($"线程名称: {currentThread.Name ?? "<未命名>"}");
                info.AppendLine($"单元状态: {currentThread.GetApartmentState()}");
                info.AppendLine($"是否后台线程: {currentThread.IsBackground}");
                info.AppendLine($"是否线程池线程: {currentThread.IsThreadPoolThread}");
                info.AppendLine($"是否存活: {currentThread.IsAlive}");
                
                // 检查是否为STA线程
                info.AppendLine($"是否STA线程: {IsSTAThread()}");
                
                // 检查UI线程可用性
                info.AppendLine($"UI线程可用: {IsUIThreadAvailable()}");
                
                // 检查WinForms应用程序状态
                try
                {
                    info.AppendLine($"WinForms窗体数量: {Application.OpenForms.Count}");
                    if (Application.OpenForms.Count > 0)
                    {
                        var mainForm = Application.OpenForms[0];
                        info.AppendLine($"主窗体: {mainForm.GetType().Name}");
                        info.AppendLine($"主窗体句柄: {mainForm.Handle}");
                        info.AppendLine($"需要调用: {mainForm.InvokeRequired}");
                    }
                }
                catch (Exception ex)
                {
                    info.AppendLine($"WinForms状态检查失败: {ex.Message}");
                }
                
                info.AppendLine($"时间戳: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
                info.AppendLine("==================");
            }
            catch (Exception ex)
            {
                info.AppendLine($"诊断信息收集失败: {ex.Message}");
            }
            
            return info.ToString();
        }

        /// <summary>
        /// 检查当前环境是否适合运行文件对话框
        /// </summary>
        /// <returns>环境检查结果</returns>
        public static FileDialogEnvironmentCheck CheckFileDialogEnvironment()
        {
            var result = new FileDialogEnvironmentCheck
            {
                Timestamp = DateTime.Now,
                ThreadId = Thread.CurrentThread.ManagedThreadId,
                IsSTAThread = IsSTAThread(),
                IsUIThreadAvailable = IsUIThreadAvailable()
            };

            // 检查具体问题
            if (!result.IsSTAThread)
            {
                result.Issues.Add("当前线程不是STA模式，文件对话框可能无法正常工作");
                result.Suggestions.Add("需要在STA线程上执行文件对话框操作");
            }

            if (!result.IsUIThreadAvailable)
            {
                result.Issues.Add("UI线程不可用，无法显示文件对话框");
                result.Suggestions.Add("确保应用程序已正确初始化WinForms环境");
            }

            // 检查应用程序入口点
            try
            {
                var entryAssembly = System.Reflection.Assembly.GetEntryAssembly();
                if (entryAssembly != null)
                {
                    var entryPoint = entryAssembly.EntryPoint;
                    if (entryPoint != null)
                    {
                        var staAttribute = entryPoint.GetCustomAttributes(typeof(STAThreadAttribute), false);
                        result.HasSTAThreadAttribute = staAttribute.Length > 0;
                        
                        if (!result.HasSTAThreadAttribute)
                        {
                            result.Issues.Add("应用程序入口点缺少STAThreadAttribute标记");
                            result.Suggestions.Add("在Main方法上添加[STAThread]属性");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result.Issues.Add($"无法检查STAThreadAttribute: {ex.Message}");
            }

            result.CanExecuteFileDialog = result.IsSTAThread && result.IsUIThreadAvailable;
            
            return result;
        }
    }
}