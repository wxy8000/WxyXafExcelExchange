# Wxy.Xaf.ExcelExchange

强大的 DevExpress XAF Excel 导入/导出模块,让数据导入导出变得简单高效!

![NuGet](https://img.shields.io/nuget/v/Wxy.Xaf.ExcelExchange)
![License](https://img.shields.io/badge/license-MIT-blue)

## ✨ 功能特性

- ✅ **一键导入导出** - 列表视图自动添加导入/导出按钮
- ✅ **批量操作** - 支持批量导出多个列表的数据
- ✅ **智能映射** - 支持自定义字段映射和列顺序
- ✅ **数据验证** - 导入时自动验证数据完整性
- ✅ **编码检测** - 自动检测 Excel 文件编码
- ✅ **跨平台** - 完美支持 WinForms 和 Blazor
- ✅ **美化工单** - 导入完成时显示固定大小的美化对话框
- ✅ **详细反馈** - 导入结果可复制,包含成功/失败/跳过统计

## 📦 安装

### 方式一:通过 NuGet 安装 (推荐)

```bash
dotnet add package Wxy.Xaf.ExcelExchange
```

或在 Visual Studio 中:
```
Tools → NuGet Package Manager → Package Manager Console
Install-Package Wxy.Xaf.ExcelExchange
```

### 方式二:项目引用

如果您有源代码,可以直接添加项目引用:
```xml
<ProjectReference Include="..\WxyXafExcelExchange\Wxy.Xaf.ExcelExchange.csproj" />
```

## 🚀 快速开始

### 步骤 1: 注册模块

#### Blazor 应用 (Startup.cs 或 Program.cs)

```csharp
// builder.Modules.Add<Wxy.Xaf.ExcelExchange.WxyXafExcelExchangeModule>();
// 或者使用完整命名空间
builder.Modules.Add<Wxy.Xaf.ExcelExchange.WxyXafExcelExchangeModule>();
```

#### WinForms 应用 (Program.cs 或 Startup.cs)

```csharp
builder.Modules.Add<Wxy.Xaf.ExcelExchange.WxyXafExcelExchangeModule>();
```

### 步骤 2: 标记业务对象

使用 `[ExcelImportExport]` 特性标记需要导入导出的业务类:

```csharp
using DevExpress.Xpo;
using Wxy.Xaf.ExcelExchange.Attributes;

namespace YourNamespace.BusinessObjects
{
    [ExcelImportExport]
    public class Product : BaseObject
    {
        [ExcelField("产品名称", SortOrder = 1)]
        public string Name { get; set; }

        [ExcelField("价格", SortOrder = 2)]
        public decimal Price { get; set; }

        [ExcelField("数量", SortOrder = 3)]
        public int Quantity { get; set; }

        [ExcelField("创建日期", SortOrder = 4, DataFormat = "yyyy-mm-dd")]
        public DateTime CreatedDate { get; set; }
    }
}
```

就这么简单!现在运行您的应用程序,您会在 Product 的列表视图看到"导入"和"导出"按钮。

## 📖 详细用法

### 1. 基础字段映射

使用 `[ExcelField]` 特性控制 Excel 列的显示名称和顺序:

```csharp
public class Employee : BaseObject
{
    [ExcelField("员工姓名", SortOrder = 1)]
    public string Name { get; set; }

    [ExcelField("工号", SortOrder = 2)]
    public string EmployeeCode { get; set; }

    [ExcelField("入职日期", SortOrder = 3, DataFormat = "yyyy-mm-dd")]
    public DateTime HireDate { get; set; }

    [ExcelField("是否在职", SortOrder = 4)]
    public bool IsActive { get; set; }
}
```

### 2. 列顺序控制

有三种方式控制 Excel 导出的列顺序:

#### 方式一:使用 SortOrder (推荐)

```csharp
[ExcelField("名称", SortOrder = 1)]
public string Name { get; set; }

[ExcelField("价格", SortOrder = 2)]
public decimal Price { get; set; }

[ExcelField("库存", SortOrder = 3)]
public int Stock { get; set; }
```

#### 方式二:使用 ColumnIndex

```csharp
[ExcelField("名称", ColumnIndex = 0)]
public string Name { get; set; }

[ExcelField("价格", ColumnIndex = 1)]
public decimal Price { get; set; }
```

#### 方式三:按属性声明顺序

如果不设置 `SortOrder` 或 `ColumnIndex`,列将按照 XPO 对象属性声明顺序输出。

### 3. 日期格式化

使用 `DataFormat` 参数指定日期格式:

```csharp
[ExcelField("出生日期", SortOrder = 1, DataFormat = "yyyy-mm-dd")]
public DateTime BirthDate { get; set; }

[ExcelField("入职时间", SortOrder = 2, DataFormat = "yyyy-mm-dd hh:mm:ss")]
public DateTime HireDateTime { get; set; }
```

支持的日期格式:
- `yyyy-mm-dd` - 2024-01-15
- `yyyy-mm-dd hh:mm:ss` - 2024-01-15 14:30:00
- `yyyy年mm月dd日` - 2024年01月15日

### 4. 数据验证

结合 XAF 的验证特性确保数据质量:

```csharp
public class Product : BaseObject
{
    [ExcelField("产品名称", SortOrder = 1)]
    [RuleRequiredField(DefaultContexts.Save)]
    public string Name { get; set; }

    [ExcelField("价格", SortOrder = 2)]
    [RuleRange(0.01, 999999)]
    public decimal Price { get; set; }

    [ExcelField("库存数量", SortOrder = 3)]
    [RuleRange(0, int.MaxValue)]
    public int Stock { get; set; }

    [ExcelField("邮箱", SortOrder = 4)]
    [RuleRequiredField(DefaultContexts.Save)]
    [RuleRegularExpression("^[\\w-\\.]+@([\\w-]+\\.)+[\\w-]{2,4}$")]
    public string Email { get; set; }
}
```

### 5. 复杂类型处理

#### 关联对象

```csharp
public class Order : BaseObject
{
    [ExcelField("订单号", SortOrder = 1)]
    public string OrderNumber { get; set; }

    [ExcelField("客户名称", SortOrder = 2)]
    [Association("Order-Customer")]
    public Customer Customer { get; set; }

    [ExcelField("总金额", SortOrder = 3)]
    public decimal TotalAmount { get; set; }
}
```

#### 枚举类型

```csharp
public enum OrderStatus
{
    [XafDisplayName("待处理")]
    Pending = 0,
    [XafDisplayName("处理中")]
    Processing = 1,
    [XafDisplayName("已完成")]
    Completed = 2,
    [XafDisplayName("已取消")]
    Cancelled = 3
}

public class Order : BaseObject
{
    [ExcelField("订单状态", SortOrder = 4)]
    public OrderStatus Status { get; set; }
}
```

## 🎨 导入对话框

模块会为导入完成显示一个固定大小(680px)的美化对话框:

### Blazor 版本特性
- 🎨 渐变紫色标题栏
- ✨ 滑入动画效果
- 📋 可复制导入结果
- 🎯 固定宽度,不会随内容变化

### WinForms 版本特性
- 🎨 渐变紫色标题栏 (#667eea)
- 📏 固定大小 680x500
- 📋 复制按钮带反馈动画
- 🎯 使用 Microsoft YaHei UI 字体

## 📊 导入结果说明

导入完成后会显示详细的统计信息:

```
✅ 导入完成!

导入统计:
- 成功: 145 条
- 失败: 2 条
- 跳过: 8 条
- 总计: 155 条

失败原因:
第 15 行: 价格格式错误
第 42 行: 必填字段为空

提示: 点击"复制"按钮可以复制完整信息
```

## 🔧 高级配置

### 自定义导入逻辑

如果需要自定义导入逻辑,可以继承 `ExcelImportExportControllerBase`:

```csharp
using Wxy.Xaf.ExcelExchange.Controllers;

public class CustomExcelImportController : ExcelImportExportControllerBase
{
    protected override void OnBeforeImport(object obj, Dictionary<string, object> rowData)
    {
        base.OnBeforeImport(obj, rowData);

        // 自定义逻辑:例如设置默认值
        if (obj is Product product)
        {
            product.CreatedDate = DateTime.Now;
            product.CreatedBy = SecuritySystem.CurrentUserName;
        }
    }
}
```

### 禁用特定类的导入导出

如果不希望某个类支持导入导出,不添加 `[ExcelImportExport]` 特性即可。

或者使用条件逻辑:

```csharp
[ExcelImportExport(Enabled = false)]
public class InternalData : BaseObject
{
    // 此类不会显示导入/导出按钮
}
```

## 🛠️ 故障排除

### 问题 1: 导入/导出按钮不显示

**原因**: 类没有添加 `[ExcelImportExport]` 特性

**解决**: 确保业务类上添加了特性:
```csharp
[ExcelImportExport]
public class YourClass : BaseObject
{
    // ...
}
```

### 问题 2: 导入时编码错误

**原因**: Excel 文件编码不正确

**解决**:
1. 使用 UTF-8 编码保存 CSV 文件
2. 或使用 .xlsx 格式(自动检测编码)

### 问题 3: 某些列无法导入

**原因**: Excel 列名与 `[ExcelField]` 特性不匹配

**解决**: 确保 Excel 列名与 `ExcelField` 参数完全一致(区分大小写)

### 问题 4: 日期显示为数字

**原因**: Excel 日期格式问题

**解决**: 在 `ExcelField` 中指定 `DataFormat`:
```csharp
[ExcelField("日期", DataFormat = "yyyy-mm-dd")]
public DateTime Date { get; set; }
```

## 📝 示例项目

完整示例项目请访问:
- **GitHub**: https://github.com/wxy8000/WxyXafExcelExchange
- **文档**: https://github.com/wxy8000/WxyXafExcelExchange/wiki

## 🌐 平台支持

| 平台 | 支持状态 | 备注 |
|:---:|:---:|:---|
| .NET 8.0 | ✅ | 推荐 |
| .NET 7.0 | ✅ | |
| DevExpress XAF 25.1.* | ✅ | |
| Blazor Server | ✅ | 完整支持 |
| WinForms | ✅ | 完整支持 |

## 🤝 贡献

欢迎提交 Issue 和 Pull Request!

1. Fork 本项目
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

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

- 查看 [完整赞助名单](https://github.com/wxy8000/WxyXafExcelExchange/blob/main/SPONSORS.md)

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情

## 📧 联系方式

- 作者: wxy8000
- GitHub: [@wxy8000](https://github.com/wxy8000)
- Email: (通过 GitHub 联系)

---

**享受使用 Wxy.Xaf.ExcelExchange!** 🚀
