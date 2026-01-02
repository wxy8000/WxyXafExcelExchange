using System;
using System.Collections.Generic;
using System.Threading;

namespace Wxy.Xaf.ExcelExchange.Threading
{
    /// <summary>
    /// 线程诊断信息
    /// </summary>
    public class ThreadDiagnosticInfo
    {
        /// <summary>
        /// 线程ID
        /// </summary>
        public int ThreadId { get; set; }

        /// <summary>
        /// 单元状态（STA/MTA/Unknown）
        /// </summary>
        public ApartmentState ApartmentState { get; set; }

        /// <summary>
        /// 是否为后台线程
        /// </summary>
        public bool IsBackground { get; set; }

        /// <summary>
        /// 是否为UI线程
        /// </summary>
        public bool IsUIThread { get; set; }

        /// <summary>
        /// 线程名称
        /// </summary>
        public string ThreadName { get; set; }

        /// <summary>
        /// 是否为线程池线程
        /// </summary>
        public bool IsThreadPoolThread { get; set; }

        /// <summary>
        /// 是否存活
        /// </summary>
        public bool IsAlive { get; set; }

        /// <summary>
        /// 时间戳
        /// </summary>
        public DateTime Timestamp { get; set; }

        /// <summary>
        /// 创建当前线程的诊断信息
        /// </summary>
        /// <returns>当前线程的诊断信息</returns>
        public static ThreadDiagnosticInfo CreateForCurrentThread()
        {
            var currentThread = Thread.CurrentThread;
            
            return new ThreadDiagnosticInfo
            {
                ThreadId = currentThread.ManagedThreadId,
                ApartmentState = currentThread.GetApartmentState(),
                IsBackground = currentThread.IsBackground,
                IsUIThread = ThreadStateChecker.IsUIThreadAvailable() && ThreadStateChecker.IsSTAThread(),
                ThreadName = currentThread.Name,
                IsThreadPoolThread = currentThread.IsThreadPoolThread,
                IsAlive = currentThread.IsAlive,
                Timestamp = DateTime.Now
            };
        }

        /// <summary>
        /// 格式化输出诊断信息
        /// </summary>
        /// <returns>格式化的诊断信息字符串</returns>
        public override string ToString()
        {
            return $"Thread[{ThreadId}] {ThreadName ?? "<unnamed>"} - {ApartmentState} " +
                   $"(Background: {IsBackground}, UI: {IsUIThread}, Pool: {IsThreadPoolThread}, " +
                   $"Alive: {IsAlive}) @ {Timestamp:HH:mm:ss.fff}";
        }

        /// <summary>
        /// 获取详细的格式化信息
        /// </summary>
        /// <returns>详细的诊断信息</returns>
        public string ToDetailedString()
        {
            var details = new System.Text.StringBuilder();
            details.AppendLine($"线程ID: {ThreadId}");
            details.AppendLine($"线程名称: {ThreadName ?? "<未命名>"}");
            details.AppendLine($"单元状态: {ApartmentState}");
            details.AppendLine($"是否后台线程: {IsBackground}");
            details.AppendLine($"是否UI线程: {IsUIThread}");
            details.AppendLine($"是否线程池线程: {IsThreadPoolThread}");
            details.AppendLine($"是否存活: {IsAlive}");
            details.AppendLine($"时间戳: {Timestamp:yyyy-MM-dd HH:mm:ss.fff}");
            return details.ToString();
        }
    }

    /// <summary>
    /// 文件对话框执行上下文
    /// </summary>
    public class FileDialogExecutionContext
    {
        /// <summary>
        /// 原始调用线程信息
        /// </summary>
        public ThreadDiagnosticInfo OriginalThread { get; set; }

        /// <summary>
        /// 实际执行线程信息
        /// </summary>
        public ThreadDiagnosticInfo ExecutionThread { get; set; }

        /// <summary>
        /// 是否需要线程切换
        /// </summary>
        public bool RequiredThreadSwitch { get; set; }

        /// <summary>
        /// 执行时间
        /// </summary>
        public TimeSpan ExecutionTime { get; set; }

        /// <summary>
        /// 是否执行成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 错误消息（如果有）
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 操作类型
        /// </summary>
        public string OperationType { get; set; }

        /// <summary>
        /// 开始时间
        /// </summary>
        public DateTime StartTime { get; set; }

        /// <summary>
        /// 结束时间
        /// </summary>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// 创建新的执行上下文
        /// </summary>
        /// <param name="operationType">操作类型</param>
        /// <returns>新的执行上下文</returns>
        public static FileDialogExecutionContext Create(string operationType)
        {
            return new FileDialogExecutionContext
            {
                OperationType = operationType,
                OriginalThread = ThreadDiagnosticInfo.CreateForCurrentThread(),
                StartTime = DateTime.Now
            };
        }

        /// <summary>
        /// 标记执行开始（在目标线程上调用）
        /// </summary>
        public void MarkExecutionStart()
        {
            ExecutionThread = ThreadDiagnosticInfo.CreateForCurrentThread();
            RequiredThreadSwitch = OriginalThread.ThreadId != ExecutionThread.ThreadId;
        }

        /// <summary>
        /// 标记执行完成
        /// </summary>
        /// <param name="success">是否成功</param>
        /// <param name="errorMessage">错误消息（如果有）</param>
        public void MarkExecutionComplete(bool success, string errorMessage = null)
        {
            EndTime = DateTime.Now;
            ExecutionTime = EndTime - StartTime;
            Success = success;
            ErrorMessage = errorMessage;
        }

        /// <summary>
        /// 获取执行摘要
        /// </summary>
        /// <returns>执行摘要字符串</returns>
        public string GetExecutionSummary()
        {
            var summary = new System.Text.StringBuilder();
            summary.AppendLine($"=== 文件对话框执行摘要 ===");
            summary.AppendLine($"操作类型: {OperationType}");
            summary.AppendLine($"执行时间: {ExecutionTime.TotalMilliseconds:F2}ms");
            summary.AppendLine($"执行结果: {(Success ? "成功" : "失败")}");
            summary.AppendLine($"需要线程切换: {RequiredThreadSwitch}");
            
            if (OriginalThread != null)
            {
                summary.AppendLine($"原始线程: {OriginalThread}");
            }
            
            if (ExecutionThread != null)
            {
                summary.AppendLine($"执行线程: {ExecutionThread}");
            }
            
            if (!string.IsNullOrEmpty(ErrorMessage))
            {
                summary.AppendLine($"错误信息: {ErrorMessage}");
            }
            
            summary.AppendLine("========================");
            return summary.ToString();
        }
    }

    /// <summary>
    /// 文件对话框环境检查结果
    /// </summary>
    public class FileDialogEnvironmentCheck
    {
        /// <summary>
        /// 检查时间戳
        /// </summary>
        public DateTime Timestamp { get; set; }

        /// <summary>
        /// 当前线程ID
        /// </summary>
        public int ThreadId { get; set; }

        /// <summary>
        /// 是否为STA线程
        /// </summary>
        public bool IsSTAThread { get; set; }

        /// <summary>
        /// UI线程是否可用
        /// </summary>
        public bool IsUIThreadAvailable { get; set; }

        /// <summary>
        /// 是否有STAThreadAttribute标记
        /// </summary>
        public bool HasSTAThreadAttribute { get; set; }

        /// <summary>
        /// 是否可以执行文件对话框
        /// </summary>
        public bool CanExecuteFileDialog { get; set; }

        /// <summary>
        /// 发现的问题列表
        /// </summary>
        public List<string> Issues { get; set; } = new List<string>();

        /// <summary>
        /// 解决建议列表
        /// </summary>
        public List<string> Suggestions { get; set; } = new List<string>();

        /// <summary>
        /// 获取检查报告
        /// </summary>
        /// <returns>格式化的检查报告</returns>
        public string GetReport()
        {
            var report = new System.Text.StringBuilder();
            report.AppendLine("=== 文件对话框环境检查报告 ===");
            report.AppendLine($"检查时间: {Timestamp:yyyy-MM-dd HH:mm:ss.fff}");
            report.AppendLine($"线程ID: {ThreadId}");
            report.AppendLine($"STA线程: {IsSTAThread}");
            report.AppendLine($"UI线程可用: {IsUIThreadAvailable}");
            report.AppendLine($"STAThread属性: {HasSTAThreadAttribute}");
            report.AppendLine($"可执行文件对话框: {CanExecuteFileDialog}");
            
            if (Issues.Count > 0)
            {
                report.AppendLine();
                report.AppendLine("发现的问题:");
                foreach (var issue in Issues)
                {
                    report.AppendLine($"  - {issue}");
                }
            }
            
            if (Suggestions.Count > 0)
            {
                report.AppendLine();
                report.AppendLine("解决建议:");
                foreach (var suggestion in Suggestions)
                {
                    report.AppendLine($"  - {suggestion}");
                }
            }
            
            report.AppendLine("=============================");
            return report.ToString();
        }
    }
}