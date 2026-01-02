# Wxy.Xaf.DataDictionary

灵活的 DevExpress XAF 数据字典管理模块,轻松管理 Lookup 数据!

![NuGet](https://img.shields.io/nuget/v/Wxy.Xaf.DataDictionary)
![License](https://img.shields.io/badge/license-MIT-blue)

## ✨ 功能特性

- ✅ **动态创建** - 运行时动态创建和管理数据字典
- ✅ **自动过滤** - 根据属性名称自动过滤相关数据字典项
- ✅ **属性级关联** - 支持属性级别的数据字典关联
- ✅ **Excel 集成** - 无缝集成 Excel 导入/导出功能
- ✅ **UI 优化** - 列表视图自动禁用数据字典链接,避免误操作
- ✅ **自动初始化** - 首次运行自动创建示例数据字典
- ✅ **智能修复** - 自动检测并修复 Order 属性值
- ✅ **跨平台** - 完美支持 WinForms 和 Blazor

## 📦 安装

### 方式一:通过 NuGet 安装 (推荐)

```bash
dotnet add package Wxy.Xaf.DataDictionary
```

或在 Visual Studio 中:
```
Tools → NuGet Package Manager → Package Manager Console
Install-Package Wxy.Xaf.DataDictionary
```

### 方式二:项目引用

```xml
<ProjectReference Include="..\WxyXafDataDictionary\Wxy.Xaf.DataDictionary.csproj" />
```

## 🚀 快速开始

### 步骤 1: 注册模块

#### Blazor 应用 (Startup.cs 或 Program.cs)

```csharp
builder.Modules.Add<Wxy.Xaf.DataDictionary.WxyXafDataDictionaryModule>();
```

#### WinForms 应用 (Program.cs 或 Startup.cs)

```csharp
builder.Modules.Add<Wxy.Xaf.DataDictionary.WxyXafDataDictionaryModule>();
```

### 步骤 2: 标记业务对象

使用 `[DataDictionary]` 特性标记需要数据字典的属性:

```csharp
using DevExpress.Xpo;
using Wxy.Xaf.DataDictionary.Attributes;

namespace YourNamespace.BusinessObjects
{
    public class Employee : BaseObject
    {
        [DataDictionary("性别")]
        public XPCollection<DataDictionaryItem> Gender { get; set; }

        [DataDictionary("部门")]
        public DataDictionaryItem Department { get; set; }

        [DataDictionary("职位")]
        public DataDictionaryItem Position { get; set; }
    }
}
```

### 步骤 3: 运行应用

首次运行时,模块会自动创建示例数据字典,您可以在应用中管理这些字典。

## 📖 详细用法

### 1. 单选数据字典

用于单选场景,例如性别、学历、状态等:

```csharp
public class Employee : BaseObject
{
    [DataDictionary("性别")]
    [EditorAttribute(typeof(DataDictionaryEditor), typeof(System.Drawing.Design.UITypeEditor))]
    public DataDictionaryItem Gender { get; set; }

    [DataDictionary("学历")]
    public DataDictionaryItem Education { get; set; }

    [DataDictionary("员工状态")]
    public DataDictionaryItem Status { get; set; }
}
```

### 2. 多选数据字典

用于多选场景,例如技能、爱好等:

```csharp
public class Employee : BaseObject
{
    [DataDictionary("技能")]
    [Association("Employee-Skills")]
    public XPCollection<DataDictionaryItem> Skills { get; set; }

    [DataDictionary("语言能力")]
    [Association("Employee-Languages")]
    public XPCollection<DataDictionaryItem> Languages { get; set; }
}
```

### 3. 数据字典自动初始化

模块会在首次运行时自动创建以下示例数据字典:

| 字典名称 | 说明 | 示例项 |
|:---|:---|:---|
| 性别 | 性别选择 | 男、女 |
| 民族 | 民族选择 | 汉族、壮族、维吾尔族等 |
| 学历 | 学历选择 | 小学、初中、高中、大专、本科、硕士、博士 |
| 政治面貌 | 政治面貌 | 群众、团员、党员 |
| 婚姻状况 | 婚姻状况 | 未婚、已婚、离异、丧偶 |
| 员工状态 | 员工状态 | 在职、离职、试用、退休 |
| 部门类型 | 部门分类 | 技术部、销售部、财务部、人事部 |

### 4. 自定义数据字典

#### 方式一:通过界面创建

1. 运行应用
2. 导航到"数据字典"(DataDictionary)
3. 点击"新建"
4. 填写字典名称
5. 在"数据字典项"(DataDictionaryItem)中添加选项

#### 方式二:通过代码创建

```csharp
using DevExpress.Xpo;

public class CustomDictionaryUpdater : ModuleUpdater
{
    public override void UpdateDatabaseAfterUpdateSchema()
    {
        base.UpdateDatabaseAfterUpdateSchema();

        // 创建"客户类型"数据字典
        var customerTypeDict = ObjectSpace.FindObject<DataDictionary>(
            CriteriaOperator.Parse("Name=?", "客户类型"));

        if (customerTypeDict == null)
        {
            customerTypeDict = ObjectSpace.CreateObject<DataDictionary>();
            customerTypeDict.Name = "客户类型";
            customerTypeDict.Save();

            // 添加数据字典项
            var items = new[] { "VIP客户", "普通客户", "潜在客户" };
            foreach (var (item, index) in items.Select((x, i) => (x, i)))
            {
                var dictItem = ObjectSpace.CreateObject<DataDictionaryItem>();
                dictItem.Name = item;
                dictItem.ParentDictionary = customerTypeDict;
                dictItem.Order = index;
                dictItem.Save();
            }
        }

        ObjectSpace.CommitChanges();
    }
}
```

### 5. 使用数据字典

#### 在 Blazor 视图中

数据字典会自动显示为下拉选择框:

```csharp
public class Customer : BaseObject
{
    public string Name { get; set; }

    [DataDictionary("客户类型")]
    public DataDictionaryItem CustomerType { get; set; }

    [DataDictionary("客户等级")]
    public DataDictionaryItem CustomerLevel { get; set; }
}
```

#### 在列表视图中

数据字典会显示对应的名称,而不是 ID:

| 姓名 | 客户类型 | 客户等级 |
|:---|:---|:---|
| 张三 | VIP客户 | A级 |
| 李四 | 普通客户 | B级 |

### 6. Order 属性自动修复

模块会自动检测并修复数据字典项的 Order 值:

```csharp
// 如果 Order 值出现断层或重复,会自动重新计算:
// 0, 1, 2, 5, 8 → 0, 1, 2, 3, 4
// 10, 10, 10 → 0, 1, 2
```

## 🎨 UI 优化

### 列表视图链接禁用

为了避免用户误点击数据字典链接跳转到详情页,模块会自动在列表视图的表格中禁用数据字典链接。

**禁用范围**:
- ✅ 列表视图表格内的链接
- ✅ 数据网格(DataGrid)中的链接

**不禁用范围**:
- ✅ 左侧导航栏
- ✅ 详情视图
- ✅ 其他非列表视图区域

### Blazor 平台实现

使用精确的 CSS 选择器,只在表格单元格内禁用链接:

```css
/* 只禁用表格内的数据字典链接 */
.main-content table a[href*='DataDictionary'],
.list-view table a[href*='DataDictionary'],
tbody tr a[href*='DataDictionary'] {
    pointer-events: none !important;
    color: inherit !important;
    text-decoration: none !important;
}

/* 明确不禁用导航栏 */
.nav-menu a[href*='DataDictionary'],
.sidebar a[href*='DataDictionary'] {
    pointer-events: auto !important;
}
```

## 📊 Excel 集成

模块集成了 Excel 导入/导出功能(需要 Wxy.Xaf.ExcelExchange):

### 导出数据字典

1. 打开数据字典列表视图
2. 点击"导出"按钮
3. 选择保存位置
4. 数据字典将以 Excel 格式导出

### 导入数据字典

1. 准备 Excel 文件,包含以下列:
   - Name (名称)
   - Order (顺序,可选)
   - ParentDictionary (父字典名称)

2. 打开数据字典列表视图
3. 点击"导入"按钮
4. 选择 Excel 文件
5. 系统会自动创建或更新数据字典

### Excel 文件格式示例

| Name | Order | ParentDictionary |
|:---|:---:|:---|
| 男 | 0 | 性别 |
| 女 | 1 | 性别 |
| 本科 | 0 | 学历 |
| 硕士 | 1 | 学历 |

## 🔧 高级配置

### 自定义数据字典编辑器

```csharp
using Wxy.Xaf.DataDictionary.Editors;

public class CustomDataDictionaryEditor : DataDictionaryEditor
{
    protected override void OnEditValueChanged()
    {
        base.OnEditValueChanged();

        // 自定义逻辑:例如联动其他字段
        if (PropertyValue is DataDictionaryItem item)
        {
            // 当选择某个数据字典项时,执行特定逻辑
        }
    }
}
```

### 过滤数据字典

```csharp
public class Employee : BaseObject
{
    // 只显示"状态"为"在职"的部门
    [DataDictionary("部门", FilterCriteria = "Status = '在职'")]
    public DataDictionaryItem Department { get; set; }
}
```

### 数据字典验证

```csharp
public class Employee : BaseObject
{
    [DataDictionary("职位")]
    [RuleRequiredField(DefaultContexts.Save)]
    public DataDictionaryItem Position { get; set; }
}
```

## 🛠️ 故障排除

### 问题 1: 数据字典不显示

**原因**:
1. 模块未正确注册
2. 数据库未初始化

**解决**:
1. 确保模块已注册: `builder.Modules.Add<WxyXafDataDictionaryModule>()`
2. 检查数据库是否已创建数据字典表

### 问题 2: 下拉列表为空

**原因**: 数据字典未创建或没有数据字典项

**解决**:
1. 导航到"数据字典"管理界面
2. 创建对应名称的数据字典
3. 添加数据字典项

### 问题 3: Order 顺序不正确

**原因**: 数据字典项的 Order 值未设置或重复

**解决**: 模块会自动修复,也可以手动设置:
```csharp
var item = ObjectSpace.CreateObject<DataDictionaryItem>();
item.Order = 0; // 设置正确的顺序
item.Save();
```

### 问题 4: 数据字典链接无法点击

**原因**: 这是预期行为,列表视图中的数据字典链接被故意禁用

**解决**: 如果需要编辑数据字典:
1. 直接在列表视图中编辑
2. 或通过左侧导航进入数据字典管理界面

## 📝 示例项目

完整示例项目请访问:
- **GitHub**: https://github.com/wxy8000/WxyXafDataDictionary
- **文档**: https://github.com/wxy8000/WxyXafDataDictionary/wiki

## 🌐 平台支持

| 平台 | 支持状态 | 备注 |
|:---:|:---:|:---|
| .NET 8.0 | ✅ | 推荐 |
| .NET 7.0 | ✅ | |
| DevExpress XAF 25.1.* | ✅ | |
| Blazor Server | ✅ | 完整支持 |
| WinForms | ✅ | 完整支持 |

## 📦 依赖项

- **Wxy.Xaf.ExcelExchange** (可选,用于 Excel 导入/导出)

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

**享受使用 Wxy.Xaf.DataDictionary!** 🚀
