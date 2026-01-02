using System;
using System.Runtime.Serialization;

namespace Wxy.Xaf.ExcelExchange.Threading
{
    /// <summary>
    /// STA线程相关异常
    /// </summary>
    [Serializable]
    public class STAThreadException : InvalidOperationException
    {
        /// <summary>
        /// 线程诊断信息
        /// </summary>
        public ThreadDiagnosticInfo ThreadInfo { get; }

        /// <summary>
        /// 环境检查结果
        /// </summary>
        public FileDialogEnvironmentCheck EnvironmentCheck { get; }

        /// <summary>
        /// 建议的解决方案
        /// </summary>
        public string SuggestedSolution { get; }

        /// <summary>
        /// 创建STAThreadException实例
        /// </summary>
        /// <param name="message">异常消息</param>
        public STAThreadException(string message) 
            : base(message)
        {
            ThreadInfo = ThreadDiagnosticInfo.CreateForCurrentThread();
            EnvironmentCheck = ThreadStateChecker.CheckFileDialogEnvironment();
            SuggestedSolution = GenerateSolution();
        }

        /// <summary>
        /// 创建STAThreadException实例
        /// </summary>
        /// <param name="message">异常消息</param>
        /// <param name="innerException">内部异常</param>
        public STAThreadException(string message, Exception innerException) 
            : base(message, innerException)
        {
            ThreadInfo = ThreadDiagnosticInfo.CreateForCurrentThread();
            EnvironmentCheck = ThreadStateChecker.CheckFileDialogEnvironment();
            SuggestedSolution = GenerateSolution();
        }

        /// <summary>
        /// 创建STAThreadException实例
        /// </summary>
        /// <param name="message">异常消息</param>
        /// <param name="threadInfo">线程诊断信息</param>
        /// <param name="environmentCheck">环境检查结果</param>
        public STAThreadException(string message, ThreadDiagnosticInfo threadInfo, FileDialogEnvironmentCheck environmentCheck) 
            : base(message)
        {
            ThreadInfo = threadInfo ?? ThreadDiagnosticInfo.CreateForCurrentThread();
            EnvironmentCheck = environmentCheck ?? ThreadStateChecker.CheckFileDialogEnvironment();
            SuggestedSolution = GenerateSolution();
        }

        /// <summary>
        /// 序列化构造函数
        /// </summary>
        /// <param name="info">序列化信息</param>
        /// <param name="context">流上下文</param>
        protected STAThreadException(SerializationInfo info, StreamingContext context) 
            : base(info, context)
        {
            // 注意：由于ThreadDiagnosticInfo和FileDialogEnvironmentCheck可能包含不可序列化的信息，
            // 这里我们重新创建它们而不是从序列化数据中恢复
            ThreadInfo = ThreadDiagnosticInfo.CreateForCurrentThread();
            EnvironmentCheck = ThreadStateChecker.CheckFileDialogEnvironment();
            SuggestedSolution = GenerateSolution();
        }

        /// <summary>
        /// 生成解决方案建议
        /// </summary>
        /// <returns>解决方案建议字符串</returns>
        private string GenerateSolution()
        {
            var solution = new System.Text.StringBuilder();
            solution.AppendLine("建议的解决方案:");

            if (EnvironmentCheck != null)
            {
                if (!EnvironmentCheck.IsSTAThread)
                {
                    solution.AppendLine("1. 确保文件对话框在STA线程上执行");
                    solution.AppendLine("   - 使用Control.Invoke()调度到UI线程");
                    solution.AppendLine("   - 或者使用UIThreadDispatcher.InvokeOnUIThreadAsync()");
                }

                if (!EnvironmentCheck.HasSTAThreadAttribute)
                {
                    solution.AppendLine("2. 在应用程序入口点添加STAThread属性");
                    solution.AppendLine("   - 在Main方法上添加[STAThread]属性");
                    solution.AppendLine("   - 例如: [STAThread] static void Main() { ... }");
                }

                if (!EnvironmentCheck.IsUIThreadAvailable)
                {
                    solution.AppendLine("3. 确保WinForms应用程序正确初始化");
                    solution.AppendLine("   - 调用Application.EnableVisualStyles()");
                    solution.AppendLine("   - 调用Application.SetCompatibleTextRenderingDefault()");
                    solution.AppendLine("   - 确保有至少一个窗体被创建");
                }

                // 添加环境检查中的具体建议
                foreach (var suggestion in EnvironmentCheck.Suggestions)
                {
                    solution.AppendLine($"   - {suggestion}");
                }
            }

            solution.AppendLine("4. 如果问题持续存在，请检查:");
            solution.AppendLine("   - 是否在正确的应用程序类型中使用（WinForms vs Console）");
            solution.AppendLine("   - 是否有其他组件修改了线程的单元状态");
            solution.AppendLine("   - 是否在异步上下文中正确处理了线程切换");

            return solution.ToString();
        }

        /// <summary>
        /// 获取详细的异常信息
        /// </summary>
        /// <returns>详细的异常信息字符串</returns>
        public string GetDetailedMessage()
        {
            var details = new System.Text.StringBuilder();
            details.AppendLine("=== STA线程异常详细信息 ===");
            details.AppendLine($"异常消息: {Message}");
            details.AppendLine();

            if (ThreadInfo != null)
            {
                details.AppendLine("当前线程信息:");
                details.AppendLine(ThreadInfo.ToDetailedString());
                details.AppendLine();
            }

            if (EnvironmentCheck != null)
            {
                details.AppendLine(EnvironmentCheck.GetReport());
                details.AppendLine();
            }

            details.AppendLine(SuggestedSolution);

            if (InnerException != null)
            {
                details.AppendLine("内部异常:");
                details.AppendLine(InnerException.ToString());
            }

            details.AppendLine("========================");
            return details.ToString();
        }

        /// <summary>
        /// 重写ToString方法以提供更详细的信息
        /// </summary>
        /// <returns>异常的字符串表示</returns>
        public override string ToString()
        {
            return GetDetailedMessage();
        }

        /// <summary>
        /// 创建一个标准的STA线程异常
        /// </summary>
        /// <param name="operation">操作名称</param>
        /// <returns>STAThreadException实例</returns>
        public static STAThreadException CreateForFileDialog(string operation)
        {
            var message = $"无法在当前线程上执行{operation}操作。文件对话框需要在STA（Single Threaded Apartment）模式的线程上运行。";
            return new STAThreadException(message);
        }

        /// <summary>
        /// 创建一个包装其他异常的STA线程异常
        /// </summary>
        /// <param name="operation">操作名称</param>
        /// <param name="innerException">内部异常</param>
        /// <returns>STAThreadException实例</returns>
        public static STAThreadException CreateForFileDialog(string operation, Exception innerException)
        {
            var message = $"执行{operation}操作时发生STA线程相关错误。";
            return new STAThreadException(message, innerException);
        }

        /// <summary>
        /// 检查异常是否可能是STA线程相关问题
        /// </summary>
        /// <param name="exception">要检查的异常</param>
        /// <returns>如果可能是STA线程问题返回true</returns>
        public static bool IsPotentialSTAThreadIssue(Exception exception)
        {
            if (exception == null) return false;

            var message = exception.Message?.ToLowerInvariant() ?? "";
            
            return message.Contains("sta") ||
                   message.Contains("single threaded apartment") ||
                   message.Contains("ole") ||
                   message.Contains("com") ||
                   message.Contains("调用线程必须为") ||
                   message.Contains("thread must be") ||
                   message.Contains("apartment");
        }
    }
}