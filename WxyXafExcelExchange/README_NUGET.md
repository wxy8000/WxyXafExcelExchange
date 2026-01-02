# Wxy.Xaf.ExcelExchange

强大的 DevExpress XAF Excel 导入/导出模块

## 功能特性

- ✅ 列表视图导入/导出按钮
- ✅ 批量导出多个列表
- ✅ 支持自定义字段映射
- ✅ 数据验证
- ✅ 编码自动检测
- ✅ 跨平台支持 (WinForms & Blazor)
- ✅ 列顺序控制 (SortOrder/ColumnIndex)
- ✅ 未设置顺序时按 XPO 属性声明顺序输出

## 使用方法

### 1. 安装 NuGet 包

```bash
dotnet add package Wxy.Xaf.ExcelExchange
```

### 2. 注册模块

**Blazor (Startup.cs):**

```csharp
builder.Modules.Add<Wxy.Xaf.ExcelExchange.WxyXafExcelExchangeModule>();
```

**WinForms (Startup.cs):**

```csharp
builder.Modules.Add<Wxy.Xaf.ExcelExchange.WxyXafExcelExchangeModule>();
```

### 3. 标记业务对象

```csharp
[ExcelImportExport]
public class Product : BaseObject
{
    [ExcelField("产品名称", SortOrder = 1)]
    public string Name { get; set; }

    [ExcelField("价格", SortOrder = 2)]
    public decimal Price { get; set; }

    [ExcelField("数量", SortOrder = 3)]
    public int Quantity { get; set; }
}
```

## 高级特性

### 列顺序控制

使用 `SortOrder` 或 `ColumnIndex` 控制 Excel 导出列顺序:

```csharp
[ExcelField("名称", SortOrder = 1)]
public string Name { get; set; }

[ExcelField("价格", SortOrder = 2)]
public decimal Price { get; set; }
```

如果未设置 `SortOrder`,列将按照 XPO 对象属性声明顺序输出。

### 自定义字段映射

```csharp
[ExcelField("员工姓名", SortOrder = 1)]
public string Name { get; set; }

[ExcelField("工号", SortOrder = 2)]
public string EmployeeCode { get; set; }

[ExcelField("入职日期", SortOrder = 3, DataFormat = "yyyy-mm-dd")]
public DateTime HireDate { get; set; }
```

### 数据验证

```csharp
[ExcelField("年龄", SortOrder = 4, MinimumValue = 18, MaximumValue = 65)]
public int Age { get; set; }

[ExcelField("邮箱", SortOrder = 5)]
[RuleRequiredField]
public string Email { get; set; }
```

## 平台支持

- .NET 8.0
- DevExpress XAF 25.1.*
- Blazor Server
- WinForms

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
| ![微信支付](https://raw.githubusercontent.com/wxy8000/WxyXafExcelExchange/main/docs/sponsors/wechat-pay.jpg) | ![支付宝](https://raw.githubusercontent.com/wxy8000/WxyXafExcelExchange/main/docs/sponsors/alipay.jpg) |

## 💖 赞助名单

感谢以下用户对本项目的赞助!

- 查看 [完整赞助名单](SPONSORS.md)

## 许可证

MIT License
