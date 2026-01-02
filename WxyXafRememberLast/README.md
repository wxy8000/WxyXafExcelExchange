# Wxy.Xaf.RememberLast

自动记忆和填充 DevExpress XAF 属性值,提升数据录入效率!

![NuGet](https://img.shields.io/nuget/v/Wxy.Xaf.RememberLast)
![License](https://img.shields.io/badge/license-MIT-blue)

## ✨ 功能特性

- ✅ **自动记忆** - 自动记忆上次输入的值
- ✅ **智能填充** - 新对象自动填充记住的值
- ✅ **属性级控制** - 精确控制哪些属性需要记忆
- ✅ **跨类型共享** - 支持在不同类型间共享最后输入的值
- ✅ **多种数据类型** - 支持字符串、数值、日期、布尔、枚举等
- ✅ **过期机制** - 支持设置记住值的过期时间
- ✅ **跨平台** - 完美支持 WinForms 和 Blazor

## 📦 安装

### 方式一:通过 NuGet 安装 (推荐)

```bash
dotnet add package Wxy.Xaf.RememberLast
```

或在 Visual Studio 中:
```
Tools → NuGet Package Manager → Package Manager Console
Install-Package Wxy.Xaf.RememberLast
```

### 方式二:项目引用

```xml
<ProjectReference Include="..\WxyXafRememberLast\Wxy.Xaf.RememberLast.csproj" />
```

## 🚀 快速开始

### 步骤 1: 注册模块和服务

#### Blazor 应用 (Startup.cs 或 Program.cs)

```csharp
// 注册模块
builder.Modules.Add<Wxy.Xaf.RememberLast.WxyXafRememberLastModule>();

// 注册存储服务
services.AddSingleton<ILastValueStorageService, MemoryLastValueStorageService>();
```

#### WinForms 应用 (Program.cs 或 Startup.cs)

```csharp
// 注册模块
builder.Modules.Add<Wxy.Xaf.RememberLast.WxyXafRememberLastModule>();

// 注册存储服务
builder.Services.AddSingleton<ILastValueStorageService, MemoryLastValueStorageService>();
```

### 步骤 2: 标记属性

使用 `[RememberLast]` 特性标记需要记忆的属性:

```csharp
using DevExpress.Xpo;
using Wxy.Xaf.RememberLast.Attributes;

namespace YourNamespace.BusinessObjects
{
    public class Order : BaseObject
    {
        [RememberLast]
        public string CustomerName { get; set; }

        [RememberLast]
        public string ShippingAddress { get; set; }

        [RememberLast]
        public string PaymentMethod { get; set; }
    }
}
```

### 步骤 3: 运行应用

现在当您创建新对象时,标记的属性会自动填充上次输入的值!

## 📖 详细用法

### 1. 基础用法

#### 字符串类型

```csharp
public class Product : BaseObject
{
    [RememberLast]
    public string Category { get; set; }

    [RememberLast]
    public string Supplier { get; set; }

    [RememberLast]
    public string Warehouse { get; set; }
}
```

#### 数值类型

```csharp
public class Product : BaseObject
{
    [RememberLast]
    public decimal DefaultPrice { get; set; }

    [RememberLast]
    public int DefaultQuantity { get; set; }

    [RememberLast]
    public double Weight { get; set; }
}
```

#### 日期类型

```csharp
public class Event : BaseObject
{
    [RememberLast]
    public DateTime StartDate { get; set; }

    [RememberLast]
    public DateTime EndDate { get; set; }
}
```

#### 布尔类型

```csharp
public class Task : BaseObject
{
    [RememberLast]
    public bool IsUrgent { get; set; }

    [RememberLast]
    public bool RequiresApproval { get; set; }
}
```

#### 枚举类型

```csharp
public enum Priority
{
    Low = 0,
    Medium = 1,
    High = 2,
    Urgent = 3
}

public class Task : BaseObject
{
    [RememberLast]
    public Priority Priority { get; set; }
}
```

### 2. 跨类型共享

使用 `ShareAcrossTypes` 参数在不同类型间共享最后输入的值:

```csharp
public class Product : BaseObject
{
    // 记住的"供应商"值会在 Product 和 Order 之间共享
    [RememberLast(ShareAcrossTypes = true)]
    public string Supplier { get; set; }
}

public class Order : BaseObject
{
    // 记住的"供应商"值会在 Product 和 Order 之间共享
    [RememberLast(ShareAcrossTypes = true)]
    public string Supplier { get; set; }
}
```

**使用场景**:
- 不同类型的订单都使用相同的供应商
- 不同类型的文档都使用相同的创建人
- 任何需要在多个业务对象间共享的默认值

### 3. 过期机制

使用 `ExpirationMinutes` 参数设置记住值的过期时间:

```csharp
public class Session : BaseObject
{
    // 30分钟后过期
    [RememberLast(ExpirationMinutes = 30)]
    public string UserName { get; set; }

    // 1小时后过期
    [RememberLast(ExpirationMinutes = 60)]
    public string WorkStation { get; set; }

    // 1天后过期
    [RememberLast(ExpirationMinutes = 1440)]
    public string DefaultDepartment { get; set; }
}
```

**使用场景**:
- 临时会话信息(30分钟)
- 班次相关信息(8小时 = 480分钟)
- 每日默认值(1440分钟 = 24小时)

### 4. 组合使用

```csharp
public class Invoice : BaseObject
{
    // 基础用法:简单记忆
    [RememberLast]
    public string PaymentTerms { get; set; }

    // 跨类型共享 + 过期时间
    [RememberLast(ShareAcrossTypes = true, ExpirationMinutes = 480)]
    public string SalesPerson { get; set; }

    // 永久记忆(不设置过期时间)
    [RememberLast]
    public string DefaultCurrency { get; set; }

    // 不记忆
    public string InvoiceNumber { get; set; }  // 没有 [RememberLast] 特性
}
```

### 5. 自定义存储服务

默认使用内存存储 `MemoryLastValueStorageService`。如果需要持久化,可以实现自定义存储服务:

```csharp
using Wxy.Xaf.RememberLast.Services;

public class DatabaseLastValueStorageService : ILastValueStorageService
{
    private readonly IObjectSpaceProvider _objectSpaceProvider;

    public DatabaseLastValueStorageService(IObjectSpaceProvider objectSpaceProvider)
    {
        _objectSpaceProvider = objectSpaceProvider;
    }

    public void SetValue(string key, object value)
    {
        using (var objectSpace = _objectSpaceProvider.CreateObjectSpace())
        {
            var storedValue = objectSpace.FindObject<LastStoredValue>(
                CriteriaOperator.Parse("Key=?", key));

            if (storedValue == null)
            {
                storedValue = objectSpace.CreateObject<LastStoredValue>();
                storedValue.Key = key;
            }

            storedValue.Value = value;
            storedValue.LastUpdated = DateTime.Now;
            objectSpace.CommitChanges();
        }
    }

    public object GetValue(string key)
    {
        using (var objectSpace = _objectSpaceProvider.CreateObjectSpace())
        {
            var storedValue = objectSpace.FindObject<LastStoredValue>(
                CriteriaOperator.Parse("Key=?", key));

            return storedValue?.Value;
        }
    }

    public bool HasValue(string key)
    {
        using (var objectSpace = _objectSpaceProvider.CreateObjectSpace())
        {
            return objectSpace.FindObject<LastStoredValue>(
                CriteriaOperator.Parse("Key=?", key)) != null;
        }
    }
}
```

**注册自定义服务**:

```csharp
// Blazor
services.AddSingleton<ILastValueStorageService, DatabaseLastValueStorageService>();

// WinForms
builder.Services.AddSingleton<ILastValueStorageService, DatabaseLastValueStorageService>();
```

## 🎯 实际应用场景

### 场景 1: 订单录入

```csharp
public class Order : BaseObject
{
    [RememberLast]
    public string CustomerName { get; set; }

    [RememberLast(ShareAcrossTypes = true)]
    public string ShippingAddress { get; set; }

    [RememberLast]
    public string PaymentMethod { get; set; }

    [RememberLast]
    public string SalesPerson { get; set; }

    public string OrderNumber { get; set; }  // 不记忆,每次不同
    public DateTime OrderDate { get; set; }  // 不记忆,使用当前日期
}
```

**效果**:
- 第二次创建订单时,客户名称、发货地址、付款方式、销售人员自动填充
- 减少重复输入,提升录入效率

### 场景 2: 产品管理

```csharp
public class Product : BaseObject
{
    [RememberLast]
    public string Category { get; set; }

    [RememberLast(ShareAcrossTypes = true)]
    public string Supplier { get; set; }

    [RememberLast]
    public string Warehouse { get; set; }

    [RememberLast]
    public decimal DefaultPrice { get; set; }

    [RememberLast(ExpirationMinutes = 480)]  // 8小时过期
    public string BatchNumber { get; set; }
}
```

### 场景 3: 会话信息

```csharp
public class WorkSession : BaseObject
{
    [RememberLast(ExpirationMinutes = 30)]  // 30分钟过期
    public string OperatorName { get; set; }

    [RememberLast(ExpirationMinutes = 30)]
    public string WorkStation { get; set; }

    [RememberLast(ExpirationMinutes = 30)]
    public string Shift { get; set; }  // 早班/中班/晚班
}
```

## 🔧 高级配置

### 全局禁用

如果需要临时禁用所有记忆功能:

```csharp
public class CustomRememberLastController : RememberLastController
{
    protected override void OnActivated()
    {
        base.OnActivated();

        // 根据条件禁用
        if (SomeCondition)
        {
            Active["RememberLast"] = false;
        }
    }
}
```

### 条件记忆

```csharp
public class Order : BaseObject
{
    private bool ShouldRememberSupplier()
    {
        // 只有特定角色才记忆供应商
        return SecuritySystem.UserIsInRole("Sales");
    }

    [RememberLast]
    public string CustomerName { get; set; }

    // 自定义逻辑控制是否记忆
    [RememberLast(Enabled = false)]
    public string Supplier { get; set; }
}
```

## 🛠️ 故障排除

### 问题 1: 属性值没有自动填充

**原因**:
1. 模块未正确注册
2. 存储服务未注册
3. 属性没有添加 `[RememberLast]` 特性

**解决**:
1. 确保模块已注册: `builder.Modules.Add<WxyXafRememberLastModule>()`
2. 确保存储服务已注册: `services.AddSingleton<ILastValueStorageService, MemoryLastValueStorageService>()`
3. 检查属性是否有 `[RememberLast]` 特性

### 问题 2: 记住的值不正确

**原因**: 可能是跨类型共享导致的冲突

**解决**:
```csharp
// 如果不希望跨类型共享,移除 ShareAcrossTypes 参数
[RememberLast]  // 只在当前类型内记忆
public string FieldName { get; set; }
```

### 问题 3: 值过期太快

**原因**: `ExpirationMinutes` 设置太短

**解决**: 增加过期时间或移除该参数(永久记忆):
```csharp
[RememberLast(ExpirationMinutes = 1440)]  // 1天
public string FieldName { get; set; }

// 或永久记忆
[RememberLast]
public string FieldName2 { get; set; }
```

### 问题 4: 应用重启后记住的值丢失

**原因**: 默认使用内存存储,应用重启后数据丢失

**解决**: 实现自定义的持久化存储服务(参见上文"自定义存储服务")

## 📊 支持的数据类型

| 数据类型 | 支持状态 | 备注 |
|:---|:---:|:---|
| `string` | ✅ | |
| `int` | ✅ | |
| `long` | ✅ | |
| `decimal` | ✅ | |
| `double` | ✅ | |
| `float` | ✅ | |
| `bool` | ✅ | |
| `DateTime` | ✅ | |
| `enum` | ✅ | |
| `Guid` | ✅ | |
| 可序列化对象 | ✅ | 需要标记为可序列化 |

## 📝 最佳实践

### 1. 选择合适的属性

**适合记忆的属性**:
- ✅ 用户经常重复输入的值
- ✅ 变化不频繁的默认值
- ✅ 业务流程中的固定步骤

**不适合记忆的属性**:
- ❌ 每次都不同的唯一标识(订单号、序列号等)
- ❌ 必须实时计算的字段
- ❌ 敏感信息(密码、密钥等)

### 2. 合理使用过期时间

```csharp
// 临时信息:短过期时间
[RememberLast(ExpirationMinutes = 30)]
public string TemporaryField { get; set; }

// 班次信息:中等过期时间
[RememberLast(ExpirationMinutes = 480)]  // 8小时
public string ShiftInfo { get; set; }

// 长期默认值:长过期时间或不设置
[RememberLast]  // 永久记忆
public string DefaultCategory { get; set; }
```

### 3. 谨慎使用跨类型共享

只在真正需要时使用 `ShareAcrossTypes`:

```csharp
// ✅ 合理使用:多个类型确实使用相同的供应商
[RememberLast(ShareAcrossTypes = true)]
public string Supplier { get; set; }

// ❌ 不合理使用:订单号和产品编号完全不同
[RememberLast(ShareAcrossTypes = true)]
public string Number { get; set; }
```

## 📝 示例项目

完整示例项目请访问:
- **GitHub**: https://github.com/wxy8000/WxyXafRememberLast
- **文档**: https://github.com/wxy8000/WxyXafRememberLast/wiki

## 🌐 平台支持

| 平台 | 支持状态 | 备注 |
|:---:|:---:|:---|
| .NET 8.0 | ✅ | 推荐 |
| .NET 7.0 | ✅ | |
| DevExpress XAF 25.1.* | ✅ | |
| Blazor Server | ✅ | 完整支持 |
| WinForms | ✅ | 完整支持 |

## 📦 依赖项

- **Wxy.Xaf.ExcelExchange** (可选,用于 Excel 导入/导出支持)

## 🤝 贡献

欢迎提交 Issue 和 Pull Request!

## ☕ 赞助 / Sponsor

如果您觉得这个项目对您有帮助,欢迎请我喝杯咖啡!

### GitHub Sponsors (推荐)

[**❤️ Sponsor 我**](https://github.com/sponsors/wxy8000)

### 微信 / 支付宝

| 微信支付 | 支付宝 |
|:---:|:---:|
| ![微信支付](https://maas-log-prod.cn-wlcb.ufileos.com/anthropic/57b5b884-86be-459a-be5d-0777921da443/8bf2ffd0c061f36b69935d8ea7cadd19.jpg?UCloudPublicKey=TOKEN_e15ba47a-d098-4fbd-9afc-a0dcf0e4e621&Expires=1767319239&Signature=+beX/qpG2VRZaoZTDqC3LpAk8dw=) | ![支付宝](https://maas-log-prod.cn-wlcb.ufileos.com/anthropic/57b5b884-86be-459a-be5d-0777921da443/367694156888ad7489fbe809cd4da586.jpg?UCloudPublicKey=TOKEN_e15ba47a-d098-4fbd-9afc-a0dcf0e4e621&Expires=1767319239&Signature=K89KYmZ%2FFjjbL%2BaDij6BmY6EJfk%3D) |

## 💖 赞助名单

感谢以下用户对本项目的赞助!

- 查看 [完整赞助名单](SPONSORS.md)

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情

## 📧 联系方式

- 作者: wxy8000
- GitHub: [@wxy8000](https://github.com/wxy8000)

---

**享受使用 Wxy.Xaf.RememberLast!** 🚀
