using System.ComponentModel;

namespace Wxy.Xaf.DataDictionary;

/// <summary>
/// 鏍囪灞炴€т负鏁版嵁瀛楀吀灞炴€?
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class DataDictionaryAttribute : Attribute
{
    /// <summary>
    /// 鏁版嵁瀛楀吀鍚嶇О
    /// </summary>
    public string DataDictionaryName { get; }

    /// <summary>
    /// 榛樿瀛楀吀椤瑰垪琛?鏍煎紡: 鍚嶇О|缂栫爜)
    /// 绀轰緥: "鍚敤|1;绂佺敤|0;鏈煡|99"
    /// </summary>
    public string DefaultItems { get; set; }

    /// <summary>
    /// 鏄惁鍏佽鐢ㄦ埛鑷畾涔夊瓧鍏搁」(榛樿 true)
    /// </summary>
    public bool AllowCustomItems { get; set; } = true;

    /// <summary>
    /// 瀛楀吀椤规弿杩?鍙€?
    /// </summary>
    public string Description { get; set; }

    public DataDictionaryAttribute(string dataDictionaryName)
    {
        DataDictionaryName = dataDictionaryName;
    }

    /// <summary>
    /// 鍒涘缓鍖呭惈榛樿椤圭殑鏁版嵁瀛楀吀鐗规€?
    /// </summary>
    public DataDictionaryAttribute(string dataDictionaryName, string defaultItems) : this(dataDictionaryName)
    {
        DefaultItems = defaultItems;
    }
}

