using System;
using System.Threading.Tasks;

namespace Wxy.Xaf.RememberLast.Services;

/// <summary>
/// 鏈€鍚庡€煎瓨鍌ㄦ湇鍔℃帴鍙?
/// </summary>
public interface ILastValueStorageService
{
    /// <summary>
    /// 淇濆瓨灞炴€у€?
    /// </summary>
    /// <param name="objectType">瀵硅薄绫诲瀷</param>
    /// <param name="propertyName">灞炴€у悕绉?/param>
    /// <param name="value">灞炴€у€?/param>
    /// <param name="shareAcrossTypes">鏄惁璺ㄧ被鍨嬪叡浜?/param>
    System.Threading.Tasks.Task SaveLastValueAsync(Type objectType, string propertyName, object? value, bool shareAcrossTypes = false);

    /// <summary>
    /// 鑾峰彇淇濆瓨鐨勫睘鎬у€?
    /// </summary>
    /// <param name="objectType">瀵硅薄绫诲瀷</param>
    /// <param name="propertyName">灞炴€у悕绉?/param>
    /// <param name="shareAcrossTypes">鏄惁璺ㄧ被鍨嬪叡浜?/param>
    /// <returns>淇濆瓨鐨勫€?濡傛灉涓嶅瓨鍦ㄥ垯杩斿洖 null</returns>
    System.Threading.Tasks.Task<object?> GetLastValueAsync(Type objectType, string propertyName, bool shareAcrossTypes = false);

    /// <summary>
    /// 娓呴櫎淇濆瓨鐨勫睘鎬у€?
    /// </summary>
    /// <param name="objectType">瀵硅薄绫诲瀷</param>
    /// <param name="propertyName">灞炴€у悕绉?/param>
    /// <param name="shareAcrossTypes">鏄惁璺ㄧ被鍨嬪叡浜?/param>
    System.Threading.Tasks.Task ClearLastValueAsync(Type objectType, string propertyName, bool shareAcrossTypes = false);

    /// <summary>
    /// 娓呴櫎鎸囧畾绫诲瀷鐨勬墍鏈変繚瀛樺€?
    /// </summary>
    /// <param name="objectType">瀵硅薄绫诲瀷</param>
    System.Threading.Tasks.Task ClearTypeValuesAsync(Type objectType);

    /// <summary>
    /// 娓呴櫎鎵€鏈夎繃鏈熺殑淇濆瓨鍊?
    /// </summary>
    System.Threading.Tasks.Task ClearExpiredValuesAsync();
}

