using System;
using System.Collections.Concurrent;
using System.Linq;
using DevExpress.Xpo;
using DevExpress.Persistent.BaseImpl;
using Task = System.Threading.Tasks.Task;

namespace Wxy.Xaf.RememberLast.Services;

/// <summary>
/// 基于内存的最后值存储服务实现
/// </summary>
public class MemoryLastValueStorageService : ILastValueStorageService
{
    // 使用静态字段确保数据在整个应用程序生命周期中持久化
    private static readonly ConcurrentDictionary<string, StoredValue> _storage = new();
    private static readonly ConcurrentDictionary<string, DateTime> _accessTime = new();

    /// <summary>
    /// 存储的值信息
    /// </summary>
    private class StoredValue
    {
        public object? Value { get; set; }
        public DateTime StoredTime { get; set; }
        public int ExpirationDays { get; set; }
    }

    public Task SaveLastValueAsync(Type objectType, string propertyName, object? value, bool shareAcrossTypes = false)
    {
        var key = GenerateKey(objectType, propertyName, shareAcrossTypes);
        System.Diagnostics.Debug.WriteLine($"[MemoryLastValueStorage] Saving value with key: {key}, type: {objectType.FullName}, property: {propertyName}");

        // 处理 XPO 对象,只存储 Oid
        object? valueToStore = value;
        if (value is BaseObject xpoObject)
        {
            valueToStore = xpoObject.Oid;
        }

        _storage[key] = new StoredValue
        {
            Value = valueToStore,
            StoredTime = DateTime.UtcNow,
            ExpirationDays = 30 // 默认30天
        };

        _accessTime[key] = DateTime.UtcNow;

        return Task.CompletedTask;
    }

    public System.Threading.Tasks.Task<object?> GetLastValueAsync(Type objectType, string propertyName, bool shareAcrossTypes = false)
    {
        var key = GenerateKey(objectType, propertyName, shareAcrossTypes);
        System.Diagnostics.Debug.WriteLine($"[MemoryLastValueStorage] Getting value with key: {key}, type: {objectType.FullName}, property: {propertyName}");
        System.Diagnostics.Debug.WriteLine($"[MemoryLastValueStorage] All keys in storage: {string.Join(", ", _storage.Keys)}");

        if (!_storage.TryGetValue(key, out var storedValue))
        {
            System.Diagnostics.Debug.WriteLine($"[MemoryLastValueStorage] Key not found: {key}");
            return System.Threading.Tasks.Task.FromResult<object?>(null);
        }

        // 检查是否过期
        if (IsExpired(storedValue))
        {
            _storage.TryRemove(key, out _);
            _accessTime.TryRemove(key, out _);
            System.Diagnostics.Debug.WriteLine($"[MemoryLastValueStorage] Key expired: {key}");
            return System.Threading.Tasks.Task.FromResult<object?>(null);
        }

        _accessTime[key] = DateTime.UtcNow;

        System.Diagnostics.Debug.WriteLine($"[MemoryLastValueStorage] Found value for key: {key}");
        return System.Threading.Tasks.Task.FromResult(storedValue.Value);
    }

    public Task ClearLastValueAsync(Type objectType, string propertyName, bool shareAcrossTypes = false)
    {
        var key = GenerateKey(objectType, propertyName, shareAcrossTypes);
        _storage.TryRemove(key, out _);
        _accessTime.TryRemove(key, out _);

        return Task.CompletedTask;
    }

    public Task ClearTypeValuesAsync(Type objectType)
    {
        var typePrefix = objectType.FullName ?? objectType.Name;

        var keysToRemove = _storage.Keys
            .Where(key => key.StartsWith(typePrefix, StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var key in keysToRemove)
        {
            _storage.TryRemove(key, out _);
            _accessTime.TryRemove(key, out _);
        }

        return Task.CompletedTask;
    }

    public Task ClearExpiredValuesAsync()
    {
        var now = DateTime.UtcNow;
        var expiredKeys = _storage
            .Where(kvp => IsExpired(kvp.Value))
            .Select(kvp => kvp.Key)
            .ToList();

        foreach (var key in expiredKeys)
        {
            _storage.TryRemove(key, out _);
            _accessTime.TryRemove(key, out _);
        }

        return Task.CompletedTask;
    }

    /// <summary>
    /// 生成存储键
    /// </summary>
    private string GenerateKey(Type objectType, string propertyName, bool shareAcrossTypes)
    {
        if (shareAcrossTypes)
        {
            // 跨类型共享,只使用属性名
            return $"Shared_{propertyName}";
        }
        else
        {
            // 按类型存储
            var typeName = objectType.FullName ?? objectType.Name;
            return $"{typeName}_{propertyName}";
        }
    }

    /// <summary>
    /// 检查值是否过期
    /// </summary>
    private bool IsExpired(StoredValue storedValue)
    {
        if (storedValue.ExpirationDays <= 0)
        {
            return false; // 永不过期
        }

        var expirationTime = storedValue.StoredTime.AddDays(storedValue.ExpirationDays);
        return DateTime.UtcNow > expirationTime;
    }
}
