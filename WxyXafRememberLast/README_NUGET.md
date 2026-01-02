# Wxy.Xaf.RememberLast

自动记忆和填充 DevExpress XAF XPO 对象属性值的模块

## 功能特性

- ✅ 自动记忆上次输入值
- ✅ 新对象自动填充
- ✅ 属性级别控制
- ✅ 跨类型共享
- ✅ 支持多种数据类型
- ✅ 过期机制
- ✅ 跨平台支持 (WinForms & Blazor)

## 使用方法

### 1. 安装 NuGet 包

```bash
dotnet add package Wxy.Xaf.RememberLast
```

### 2. 注册模块

**Blazor (Startup.cs):**

```csharp
builder.Modules.Add<Wxy.Xaf.RememberLast.WxyXafRememberLastModule>();

// 注册存储服务
services.AddSingleton<ILastValueStorageService, MemoryLastValueStorageService>();
```

**WinForms (Startup.cs):**

```csharp
builder.Modules.Add<Wxy.Xaf.RememberLast.WxyXafRememberLastModule>();

// 注册存储服务
builder.Services.AddSingleton<ILastValueStorageService, MemoryLastValueStorageService>();
```

### 3. 标记属性

```csharp
public class Product : BaseObject
{
    [RememberLast]
    public string Category { get; set; }

    [RememberLast(ShareAcrossTypes = true)]
    public string Supplier { get; set; }

    [RememberLast(ExpirationMinutes = 30)]
    public string Warehouse { get; set; }
}
```

## 高级特性

### 跨类型共享

使用 `ShareAcrossTypes = true` 可以在不同类型之间共享最后输入的值:

```csharp
public class Product : BaseObject
{
    [RememberLast(ShareAcrossTypes = true)]
    public string Supplier { get; set; }
}

public class Order : BaseObject
{
    [RememberLast(ShareAcrossTypes = true)]
    public string Supplier { get; set; }
}
```

### 过期机制

设置 `ExpirationMinutes` 可以让记住的值在指定时间后过期:

```csharp
[RememberLast(ExpirationMinutes = 30)]
public string TemporaryField { get; set; }
```

### 自定义存储服务

实现 `ILastValueStorageService` 接口可以自定义存储方式:

```csharp
public class DatabaseLastValueStorageService : ILastValueStorageService
{
    // 实现数据库持久化
}
```

## 支持的数据类型

- string
- int, decimal, double 等数值类型
- DateTime
- bool
- enum
- 任何可序列化的对象类型

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
| ![微信支付](https://raw.githubusercontent.com/wxy8000/WxyXafExcelExchange/main/docs/sponsors/wechat-pay.jpg) | ![支付宝](https://raw.githubusercontent.com/wxy8000/WxyXafExcelExchange/main/docs/sponsors/alipay.jpg) |

## 💖 赞助名单

感谢以下用户对本项目的赞助!

- 查看 [完整赞助名单](SPONSORS.md)

## 许可证

MIT License
