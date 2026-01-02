using System;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using DevExpress.ExpressApp;
using DevExpress.ExpressApp.Actions;
using DevExpress.ExpressApp.Xpo;
using DevExpress.Xpo;
using DevExpress.Persistent.BaseImpl;
using Wxy.Xaf.RememberLast.Attributes;
using Wxy.Xaf.RememberLast.Services;

namespace Wxy.Xaf.RememberLast.Controllers;

/// <summary>
/// XPO对象属性值记忆功能控制器
/// 在对象保存时记录属性值,在新建对象时自动填充
/// </summary>
public class RememberLastObjectController : ViewController
{
    private ILastValueStorageService? _storageService;
    private BaseObject? _currentObject;
    private Session? _session;
    private bool _hasFilledValues = false;

    public RememberLastObjectController()
    {
    }

    protected override void OnActivated()
    {
        base.OnActivated();

        // 只在 DetailView 中激活
        if (View is not DetailView detailView)
        {
            return;
        }

        System.Diagnostics.Debug.WriteLine($"[RememberLastObject] Controller activated for view: {View.Id}");

        // 获取存储服务
        _storageService = Application.ServiceProvider?.GetService(typeof(ILastValueStorageService)) as ILastValueStorageService
            ?? new MemoryLastValueStorageService();

        // 监听对象视图的创建和保存事件
        if (View.ObjectSpace is XPObjectSpace objectSpace)
        {
            _session = objectSpace.Session;

            // 监听当前对象变更
            detailView.CurrentObjectChanged += DetailView_CurrentObjectChanged;

            // 监听 ObjectSpace 的 committing 事件
            View.ObjectSpace.Committing += ObjectSpace_Committing;

            // 立即检查是否需要填充值
            if (View.CurrentObject is BaseObject baseObject)
            {
                _currentObject = baseObject;
                bool isNew = View.ObjectSpace.IsNewObject(baseObject);
                System.Diagnostics.Debug.WriteLine($"[RememberLastObject] OnActivated - Current object: {baseObject.GetType().Name}, IsNew: {isNew}");

                if (isNew && !_hasFilledValues)
                {
                    // 同步填充值
                    FillRememberedValues(baseObject);
                    _hasFilledValues = true;
                }
            }
        }
    }

    private void DetailView_CurrentObjectChanged(object? sender, EventArgs e)
    {
        if (View.CurrentObject is not BaseObject baseObject || _session == null)
        {
            return;
        }

        _currentObject = baseObject;
        _hasFilledValues = false;

        // 如果是新对象,尝试填充上次保存的值
        if (View.ObjectSpace.IsNewObject(baseObject))
        {
            System.Diagnostics.Debug.WriteLine($"[RememberLastObject] CurrentObjectChanged - Filling values for new object: {baseObject.GetType().Name}");
            FillRememberedValues(baseObject);
            _hasFilledValues = true;
        }
    }

    private void ObjectSpace_Committing(object? sender, CancelEventArgs e)
    {
        // 在保存前记录当前属性值
        if (_currentObject != null)
        {
            System.Diagnostics.Debug.WriteLine($"[RememberLastObject] Committing - storing values for {_currentObject.GetType().Name}");
            StoreCurrentValues(_currentObject);
        }
    }

    protected override void OnDeactivated()
    {
        base.OnDeactivated();

        if (View is DetailView detailView)
        {
            detailView.CurrentObjectChanged -= DetailView_CurrentObjectChanged;
        }

        if (View.ObjectSpace != null)
        {
            View.ObjectSpace.Committing -= ObjectSpace_Committing;
        }

        _currentObject = null;
        _session = null;
        _hasFilledValues = false;
    }

    /// <summary>
    /// 填充记住的属性值
    /// </summary>
    private void FillRememberedValues(BaseObject obj)
    {
        if (_storageService == null || _session == null)
        {
            System.Diagnostics.Debug.WriteLine("[RememberLastObject] Storage service or session is null");
            return;
        }

        var objectType = obj.GetType();
        System.Diagnostics.Debug.WriteLine($"[RememberLastObject] Scanning properties for {objectType.Name}");

        var properties = objectType.GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanWrite && p.CanRead);

        int filledCount = 0;

        foreach (var property in properties)
        {
            // 检查是否有 RememberLast 特性
            var rememberAttr = property.GetCustomAttribute<RememberLastAttribute>();
            if (rememberAttr == null || !rememberAttr.Enabled)
            {
                continue;
            }

            System.Diagnostics.Debug.WriteLine($"[RememberLastObject] Found RememberLast attribute on property: {property.Name}");

            try
            {
                // 获取保存的值(使用ConfigureAwait避免死锁)
                var storedValue = _storageService.GetLastValueAsync(objectType, property.Name, rememberAttr.ShareAcrossTypes)
                    .GetAwaiter().GetResult();
                if (storedValue == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[RememberLastObject] No stored value for {property.Name}");
                    continue;
                }

                // 设置属性值
                if (property.PropertyType == typeof(string))
                {
                    property.SetValue(obj, storedValue?.ToString());
                    filledCount++;
                    System.Diagnostics.Debug.WriteLine($"[RememberLastObject] Filled string property {property.Name} = {storedValue}");
                }
                else if (property.PropertyType == typeof(int) || property.PropertyType == typeof(int?))
                {
                    if (int.TryParse(storedValue.ToString(), out var intVal))
                    {
                        property.SetValue(obj, intVal);
                        filledCount++;
                        System.Diagnostics.Debug.WriteLine($"[RememberLastObject] Filled int property {property.Name} = {intVal}");
                    }
                }
                else if (property.PropertyType == typeof(decimal) || property.PropertyType == typeof(decimal?))
                {
                    if (decimal.TryParse(storedValue.ToString(), out var decimalVal))
                    {
                        property.SetValue(obj, decimalVal);
                        filledCount++;
                        System.Diagnostics.Debug.WriteLine($"[RememberLastObject] Filled decimal property {property.Name} = {decimalVal}");
                    }
                }
                else if (property.PropertyType == typeof(bool) || property.PropertyType == typeof(bool?))
                {
                    if (bool.TryParse(storedValue.ToString(), out var boolVal))
                    {
                        property.SetValue(obj, boolVal);
                        filledCount++;
                        System.Diagnostics.Debug.WriteLine($"[RememberLastObject] Filled bool property {property.Name} = {boolVal}");
                    }
                }
                else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(DateTime?))
                {
                    if (DateTime.TryParse(storedValue.ToString(), out var dateVal))
                    {
                        property.SetValue(obj, dateVal);
                        filledCount++;
                        System.Diagnostics.Debug.WriteLine($"[RememberLastObject] Filled DateTime property {property.Name} = {dateVal}");
                    }
                }
                // 处理 XPO 关联对象
                else if (typeof(BaseObject).IsAssignableFrom(property.PropertyType))
                {
                    // 如果存储的是 Oid,则加载对应的对象
                    if (storedValue is Guid oid)
                    {
                        var relatedObject = _session.GetObjectByKey(property.PropertyType, oid);
                        if (relatedObject != null)
                        {
                            property.SetValue(obj, relatedObject);
                            filledCount++;
                            System.Diagnostics.Debug.WriteLine($"[RememberLastObject] Filled XPO object property {property.Name} with Oid = {oid}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录错误但继续处理其他属性
                System.Diagnostics.Debug.WriteLine($"[RememberLastObject] 填充属性 {property.Name} 失败: {ex.Message}");
            }
        }

        System.Diagnostics.Debug.WriteLine($"[RememberLastObject] Total properties filled: {filledCount}");
    }

    /// <summary>
    /// 保存当前属性值
    /// </summary>
    private void StoreCurrentValues(BaseObject obj)
    {
        if (_storageService == null)
        {
            System.Diagnostics.Debug.WriteLine("[RememberLastObject] Storage service is null, cannot store values");
            return;
        }

        var objectType = obj.GetType();
        System.Diagnostics.Debug.WriteLine($"[RememberLastObject] Storing values for {objectType.Name}");

        var properties = objectType.GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanWrite && p.CanRead);

        int storedCount = 0;

        foreach (var property in properties)
        {
            // 检查是否有 RememberLast 特性
            var rememberAttr = property.GetCustomAttribute<RememberLastAttribute>();
            if (rememberAttr == null || !rememberAttr.Enabled)
            {
                continue;
            }

            try
            {
                // 获取当前属性值
                var value = property.GetValue(obj);
                System.Diagnostics.Debug.WriteLine($"[RememberLastObject] Storing property {property.Name} = {value?.ToString() ?? "null"}");

                // 保存值(使用ConfigureAwait避免死锁)
                _storageService.SaveLastValueAsync(
                    objectType,
                    property.Name,
                    value,
                    rememberAttr.ShareAcrossTypes
                ).GetAwaiter().GetResult();
                storedCount++;
            }
            catch (Exception ex)
            {
                // 记录错误但继续处理其他属性
                System.Diagnostics.Debug.WriteLine($"[RememberLastObject] 保存属性 {property.Name} 失败: {ex.Message}");
            }
        }

        System.Diagnostics.Debug.WriteLine($"[RememberLastObject] Total properties stored: {storedCount}");
    }
}
