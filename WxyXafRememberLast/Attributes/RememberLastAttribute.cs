using System;
using System.ComponentModel;

namespace Wxy.Xaf.RememberLast.Attributes;

/// <summary>
/// 鏍囪灞炴€ч渶瑕佽浣忎笂娆¤緭鍏ョ殑鍊?
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class RememberLastAttribute : Attribute
{
    /// <summary>
    /// 鏄惁鍚敤璁颁綇涓婃鍊煎姛鑳?
    /// </summary>
    public bool Enabled { get; set; } = true;

    /// <summary>
    /// 璁板繂鐨勪紭鍏堢骇(鍊艰秺澶т紭鍏堢骇瓒婇珮)
    /// 褰撴湁澶氫釜鍏宠仈瀵硅薄鏃?浼樺厛浣跨敤楂樹紭鍏堢骇鐨勫€?
    /// </summary>
    public int Priority { get; set; } = 0;

    /// <summary>
    /// 璁板繂鐨勬渶澶ф暟閲?0琛ㄧず涓嶉檺鍒?
    /// 鐢ㄤ簬鍒楄〃绫诲瀷鐨勫睘鎬?
    /// </summary>
    public int MaxCount { get; set; } = 10;

    /// <summary>
    /// 鏄惁璺ㄧ被鍨嬪叡浜蹇嗗€?
    /// 濡傛灉涓?true,鍒欑浉鍚屽睘鎬у悕鐨勪笉鍚岀被鍨嬩細鍏变韩璁板繂鍊?
    /// </summary>
    public bool ShareAcrossTypes { get; set; } = false;

    /// <summary>
    /// 璁板繂鍊肩殑杩囨湡鏃堕棿(澶?
    /// 0 琛ㄧず姘镐笉杩囨湡
    /// </summary>
    public int ExpirationDays { get; set; } = 30;

    public RememberLastAttribute()
    {
    }

    public RememberLastAttribute(bool enabled)
    {
        Enabled = enabled;
    }
}

