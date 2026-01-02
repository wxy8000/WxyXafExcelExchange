using DevExpress.ExpressApp;
using DevExpress.ExpressApp.ApplicationBuilder;
using DevExpress.ExpressApp.Blazor;
using DevExpress.ExpressApp.Security;
using DevExpress.ExpressApp.Security.ClientServer;
using DevExpress.ExpressApp.SystemModule;
using DevExpress.ExpressApp.Updating;
using DevExpress.ExpressApp.Xpo;
using XafExcelDemo.Blazor.Server.Services;

namespace XafExcelDemo.Blazor.Server
{
    public class XafExcelDemoBlazorApplication : BlazorApplication
    {
        public XafExcelDemoBlazorApplication()
        {
            ApplicationName = "XafExcelDemo";
            CheckCompatibilityType = DevExpress.ExpressApp.CheckCompatibilityType.DatabaseSchema;
            DatabaseVersionMismatch += XafExcelDemoBlazorApplication_DatabaseVersionMismatch;
        }
        protected override void OnSetupStarted()
        {
            base.OnSetupStarted();

#if DEBUG
            if(System.Diagnostics.Debugger.IsAttached && CheckCompatibilityType == CheckCompatibilityType.DatabaseSchema) {
                DatabaseUpdateMode = DatabaseUpdateMode.UpdateDatabaseAlways;
            }
#endif
        }
        void XafExcelDemoBlazorApplication_DatabaseVersionMismatch(object sender, DatabaseVersionMismatchEventArgs e)
        {
            // 自动更新数据库,无论是否在调试模式下
            e.Updater.Update();
            e.Handled = true;
        }
    }
}
