# Wxy.Xaf.DataDictionary

灵活的 DevExpress XAF 数据字典管理模块

## 功能特性

- ✅ 动态数据字典创建
- ✅ 自动数据过滤
- ✅ 属性级关联
- ✅ Excel 导入/导出支持
- ✅ 超链接禁用(列表视图)
- ✅ 初始数据自动初始化
- ✅ Order 属性自动修复
- ✅ 跨平台支持 (WinForms & Blazor)

## 使用方法

### 1. 安装 NuGet 包

```bash
dotnet add package Wxy.Xaf.DataDictionary
```

### 2. 注册模块

**Blazor (Startup.cs):**

```csharp
builder.Modules.Add<Wxy.Xaf.DataDictionary.WxyXafDataDictionaryModule>();
```

**WinForms (Startup.cs):**

```csharp
builder.Modules.Add<Wxy.Xaf.DataDictionary.WxyXafDataDictionaryModule>();
```

### 3. 标记业务对象

```csharp
public class Product : BaseObject
{
    [DataDictionary("产品类别")]
    public XPCollection<DataDictionaryItem> Category { get; set; }

    [DataDictionary("计量单位")]
    public DataDictionaryItem Unit { get; set; }
}
```

## 高级特性

### 自动初始化

模块会在首次运行时自动创建示例数据字典:
- 性别
- 民族
- 学历
- 政治面貌
- 婚姻状况
- 员工状态
- 部门类型

### 自定义数据字典

```csharp
[DataDictionary("产品类别")]
[EditorAttribute(typeof(DataDictionaryEditor), typeof(System.Drawing.Design.UITypeEditor))]
public DataDictionaryItem Category { get; set; }
```

### Order 属性自动修复

模块会自动检测并修复数据字典项的 Order 值,确保顺序正确。

## 平台支持

- .NET 8.0
- DevExpress XAF 25.1.*
- Blazor Server
- WinForms

## 依赖项

- Wxy.Xaf.Excel (用于 Excel 导入/导出支持)

## 文档

详细文档请访问: https://github.com/wxy8000/wxyXafExcel

## ☕ 赞助 / Sponsor

如果您觉得这个项目对您有帮助,欢迎请我喝杯咖啡!

### GitHub Sponsors (推荐)

通过 GitHub Sponsors 赞助,支持项目持续发展:

[**❤️ Sponsor 我**](https://github.com/sponsors/wxy8000)

### 微信 / 支付宝

如需一次性赞助,可以通过以下方式:

| 微信支付 | 支付宝 |
|:---:|:---:|
| ![微信支付](https://raw.githubusercontent.com/wxy8000/wxyXafExcel/main/docs/sponsors/wechat-pay.jpg) | ![支付宝](https://raw.githubusercontent.com/wxy8000/wxyXafExcel/main/docs/sponsors/alipay.jpg) |

## 💖 赞助名单

感谢以下用户对本项目的赞助!

- 查看 [完整赞助名单](SPONSORS.md)

## 许可证

MIT License
