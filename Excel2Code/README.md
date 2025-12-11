# Excel2Code 使用指南

## 概述

Excel2Code 是一个自动化工具，用于将 Excel 配置表转换为 C# 代码。它支持多种数据类型、集合类型和高级特性，帮助游戏开发团队快速生成类型安全的配置代码。

## 核心功能

- **自动代码生成**：将 Excel (.xls/.xlsx) 转换为 C# 类定义和数据初始化代码
- **类型系统**：支持基础类型、数组、字典、自定义数据结构
- **增量编译**：基于文件修改时间戳的增量生成机制
- **编译验证**：自动编译生成的 C# 代码并报告错误
- **性能优化**：使用流式读取处理大型 Excel 文件

## 使用方式

### 1. 命令行参数

#### 1.1 交互式���入
```bash
dotnet run
```
运行后提示输入目录路径，然后生成该目录下所有 Excel 文件的代码。

#### 1.2 批量目录生成
```bash
dotnet run -- -dir <目录路径> -out <输出目录> -compile <编译目录> -ignore <忽略文件>
```

**参数说明：**
- `-dir <path>`：Excel 文件所在目录
- `-out <path>`：生成的 C# 文件输出目录（默认：`./exportcsharp`）
- `-compile <path>`：编译验证的目录（默认：`./exportcsharp`）
- `-ignore <names>`：忽略的文件名（逗号分隔，不含扩展名）

**示例：**
```bash
# 生成 Configs 目录下的所有 Excel，输出到 exportcsharp，忽略 test 和 temp 文件
dotnet run -- -dir ./Configs -out ./exportcsharp -compile ./exportcsharp -ignore test,temp
```

#### 1.3 单文件/多文件生成
```bash
dotnet run -- file1.xlsx file2.xlsx
```
指定具体的 Excel 文件路径进行生成。

### 2. Excel 文件规范

#### 2.1 文件命名
- **推荐格式**：`Configs.xlsx`, `Items.xlsx`, `Skills.xlsx` 等
- **自动重命名**：如果文件名不含 `Config_` 前缀，生成的类名会自动添加
  - `Items.xlsx` → `Config_Items.cs`
  - `Config_Skills.xlsx` → `Config_Skills.cs`
- **忽略规则**：
  - 以 `~` 开头或结尾的文件（临时文件）
  - 文件名包含空格的文件

#### 2.2 Sheet 命名规则

| Sheet 名格式 | 用途 | 生成代码 |
|-------------|------|----------|
| `#参数表` | 常量/静态变量定义 | `public const/static` 字段 |
| `普通名称` | 数据表（Dictionary 模式） | `Dictionary<TKey, TData>` + 查询方法 |
| `名称|类名` | 自定义类名的数据表 | 使用指定类名 |
| `类型 名称|类名` | 带类型前缀的数据表 | 生成 `类名Data` 类 |
| `=附加表` | 附加数据表 | 合并到前一个表 |
| `Sheet1`, `Sheet2` 等 | 被忽略 | 不生成代码 |

#### 2.3 表格结构（标准数据表）

Excel 表格必须按以下格式定义：

| 行号 | 内容 | 说明 |
|------|------|------|
| 第 1 行 | 中文注释 | 字段说明（会生成为注释） |
| 第 2 行 | 数据类型 | 如 `int`, `string`, `float[]`, `Dictionary<int,string>` |
| 第 3 行 | 字段名 | C# 字段名（必须符合命名规范） |
| 第 4+ 行 | 数据 | 实际配置数据 |

**示例：**

| ID | 名称 | 等级 | 属性列表 |
|---|---|---|---|
| int | string | int | int[] |
| Id | Name | Level | Attributes |
| 1001 | 新手剑 | 1 | 10,5,3 |
| 1002 | 钢铁剑 | 5 | 25,12,8 |

#### 2.4 表格结构（参数表）

参数表（Sheet 名以 `#` 开头）使用不同的布局：

| 列号 | 内容 |
|------|------|
| 第 1 列 | 数据类型（可选，留空则自动推断） |
| 第 2 列 | 变量名 |
| 第 3 列 | 值 |
| 第 4 列 | 中文注释（可选） |

**示例：**

| 类型 | 变量名 | 值 | 注释 |
|------|--------|-----|------|
| int | MaxLevel | 100 | 最大等级 |
| float | CritRate | 0.25f | 暴击率 |
| string | ServerUrl | https://api.game.com | 服务器地址 |

## 支持的数据类型

### 基础类型

| Excel 类型 | C# 类型 | 示例值 | 说明 |
|-----------|---------|--------|------|
| `int` | `int` | `100` | 整数 |
| `long` | `long` | `1000000` | 长整数 |
| `float` | `float` | `3.14` | 自动添加 `f` 后缀 |
| `string` | `string` | `文本内容` | 自动添加 `@""` 引号 |
| `bool` | `bool` | `true` / `false` | 布尔值 |

### 枚举类型

**新功能**：支持枚举类型的自动识别和处理。

| Excel 类型 | 识别规则 | 实际 C# 类型 | Excel 值示例 | 生成代码 |
|-----------|---------|-------------|--------------|----------|
| `EnumHeroType` | 以 `Enum` 开头 | `HeroType` | `HeroType.Magic` | `HeroType.Magic` |
| `HeroTypeEnum` | 以 `Enum` 结尾 | `HeroType` | `HeroType.Physical` | `HeroType.Physical` |
| `EnumItemQuality` | 以 `Enum` 开头 | `ItemQuality` | `ItemQuality.Rare` | `ItemQuality.Rare` |

**枚举数组支持：**

| Excel 类型 | 实际 C# 类型 | Excel 值示例 | 生成代码 |
|-----------|-------------|--------------|----------|
| `EnumHeroType[]` | `HeroType[]` | `HeroType.Magic,HeroType.Physical` | `new HeroType[]{HeroType.Magic,HeroType.Physical}` |

**使用说明：**
1. **类型定义**：在 Excel 的类型行中使用 `EnumXXX` 或 `XXXEnum` 格式
2. **实际类型**：工具会自动去除 `Enum` 前缀/后缀，得到真实的枚举类型名
3. **值格式**：在数据行中填写完整的枚举值（如 `HeroType.Magic`）
4. **前提条件**：确保枚举类型已在项目中定义

**示例：**

**Excel 表格（Sheet: Skills）：**

| ID | 技能名 | 英雄类型 | 品质等级 |
|---|--------|---------|---------|
| int | string | EnumHeroType | EnumItemQuality |
| Id | Name | HeroType | Quality |
| 1001 | 火球术 | HeroType.Magic | ItemQuality.Rare |
| 1002 | 斩击 | HeroType.Physical | ItemQuality.Common |

**生成代码：**
```csharp
public partial class Skills
{
    public int Id;
    public string Name;
    public HeroType HeroType;  // 自动去除了 Enum 前缀
    public ItemQuality Quality; // 自动去除了 Enum 前缀
}

private static Skills CreateSkills_1001()
{
    return new Skills()
    {
        Id = 1001,
        Name = @"火球术",
        HeroType = HeroType.Magic,
        Quality = ItemQuality.Rare,
    };
}
```

### 集合类型

#### 数组（`类型[]`）

| Excel 类型 | C# 类型 | Excel 值示例 | 生成代码 |
|-----------|---------|--------------|----------|
| `int[]` | `int[]` | `1,2,3` | `new int[]{1,2,3}` |
| `float[]` | `float[]` | `1.5,2.0` | `new float[]{1.5f,2.0f}` |
| `string[]` | `string[]` | `a,b,c` | `new string[]{"a","b","c"}` |

**二维数组支持：**
- Excel 类型：`int[][]`
- Excel 值：`[1,2],[3,4]`
- 生成代码：`new int[][]{new[]{1,2},new[]{3,4}}`

#### 字典（`Dictionary<TKey, TValue>`）

| Excel 类型 | Excel 值示例 | 生成代码 |
|-----------|--------------|----------|
| `Dictionary<int,string>` | `{1,物品A},{2,物品B}` | `new Dictionary<int,string>(){{1,"物品A"},{2,"物品B"}}` |
| `map<int,string>` | 同上 | 同上（`map` 会自动转换为 `Dictionary`） |
| `dic<int,string>` | 同上 | 同上（`dic` 会自动转换为 `Dictionary`） |

### 自定义数据结构

#### DataVector2 / DataVector3

| Excel 类型 | Excel 值 | 用途 |
|-----------|----------|------|
| `v2` | `1.0,2.0` | 二维向量 |
| `v2[]` | `1.0,2.0,3.0,4.0` | 向量数组 |
| `v3` | `1.0,2.0,3.0` | 三维向量 |
| `v3[]` | `1.0,2.0,3.0,4.0,5.0,6.0` | 向量数组 |

生成的代码会调用 `FromString` 方法：
```csharp
Position = new DataVector3().FromString("1.0,2.0,3.0")
```

#### 特殊类型

| Excel 类型 | 转换为 | 说明 |
|-----------|--------|------|
| `luatable` | `string` | Lua 表定义 |
| `luacode` | `string` | Lua 代码片段 |

### 类型过滤器

在类型定义中添加 `:` 后缀可以附加元数据（生成时会自动去除）：

```
int:id       → int
string:name  → string
```

## 生成代码结构

### 1. 参数表（Sheet 名：`#参数`）

**Excel 输入：**

| 类型 | 变量名 | 值 | 注释 |
|------|--------|-----|------|
| int | MaxLevel | 100 | 最大等级 |
| string | GameName | "My Game" | 游戏名称 |

**生成代码：**
```csharp
public const int MaxLevel = 100; /*最大等级*/
public const string GameName = @"My Game"; /*游戏名称*/
```

### 2. Dictionary 数据表（默认模式）

**Excel 输入（Sheet 名：`Items`）：**

| ID | 名称 | 价格 |
|---|---|---|
| int | string | int |
| Id | Name | Price |
| 1001 | 新手剑 | 100 |
| 1002 | 钢铁剑 | 500 |

**生成代码：**
```csharp
public partial class Items
{
    public int Id;
    public string Name;
    public int Price;
}

public static Dictionary<int, Items> dItems = new Dictionary<int, Items>();

// 查询方法（延迟加载）
public static Items OnGetFrom_dItems(int id)
{
    Items data = null;
    if (dItems.TryGetValue(id, out data))
    {
        return data;
    }

    var t = typeof(Config_Items);
    var m = t.GetMethod($"CreateItems_{id}", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
    if (m == null) return null;
    data = m.Invoke(null, null) as Items;
    dItems[id] = data;
    return data;
}

// 私有创建方法
private static Items CreateItems_1001()
{
    return new Items()
    {
        Id = 1001,
        Name = @"新手剑",
        Price = 100,
    };
}

// 索引列表
public static List<int> allItemss = new List<int>()
{
    1001,
    1002,
};

// 按索引查询
public static Items OnGetFrom_ItemsByIndex(int index)
{
    if (index < 0 || index >= allItemss.Count) return null;
    return OnGetFrom_dItems(allItemss[index]);
}
```

### 3. List 数据表（注释包含 "list"）

**Excel 输入（Sheet 名：`Levels`，第一行注释包含 "list"）：**

| 等级（list） | 经验 |
|-------------|------|
| int | int |
| Level | Exp |
| 1 | 100 |
| 2 | 300 |

**生成代码：**
```csharp
public static List<Levels> mLevels = new List<Levels>()
{
    new Levels
    {
        Level = 1,
        Exp = 100,
    },
    new Levels
    {
        Level = 2,
        Exp = 300,
    },
};
```

### 4. 附加表（Sheet 名：`=附加数据`）

以 `=` 开头的 Sheet 会合并到前一个普通 Sheet 的数据中。

## 高级特性

### 1. 预定义指令（`##` 前缀）

**新功能**：支持在 Excel 表格中使用预定义指令，通过 `##` 前缀标记的特殊行来控制代码生成行为。

#### 1.1 `##namespace` 指令

在 Sheet 的第一行或任意数据行使用 `##namespace` 指令，可以在生成的 C# 文件中自动添加 `using` 语句。

**语法：**
```
##namespace NamespaceName1, NamespaceName2, ...
```

**示例 1：Sheet 级别指令**

**Excel（Heroes.xlsx - Sheet: Heroes 第一行）：**

| ##namespace Server |
|--------------------|
| 英雄ID | 英雄名 | 职业类型 |
| int | string | EnumHeroType |
| Id | Name | Type |
| 3001 | 火焰法师 | HeroType.Magic |

**生成代码：**
```csharp
//本文件为自动生成，请勿手动修改
//
using System;
using System.Collections.Generic;
using Server;  // ← 自动添加的命名空间

public partial class Config_Heroes
{
    // ...
}
```

**示例 2：多个命名空间**

| ##namespace Server, Client, GameLogic |
|--------------------------------------|
| ... |

**生成代码：**
```csharp
using System;
using System.Collections.Generic;
using Client;       // ← 按字母顺序排列
using GameLogic;    // ← 按字母顺序排列
using Server;       // ← 按字母顺序排列
```

**示例 3：数据行中的指令**

你也可以在数据行中使用 `##namespace` 指令（会被跳过不生成数据）：

| ID | 名称 | 类型 |
|---|---|---|
| int | string | EnumItemType |
| Id | Name | Type |
| ##namespace Server.Items |
| 1001 | 新手剑 | ItemType.Weapon |

**特性说明：**
- ✅ 可以在 Sheet 的第一行使用（推荐）
- ✅ 可以在任意数据行使用
- ✅ 支持多个命名空间，用逗号分隔
- ✅ 多个 Sheet 中的 `##namespace` 会自动合并去重
- ✅ 生成的 `using` 语句会按字母顺序排列
- ✅ **智能行号递增**：工具会自动跳过 `##` 指令行，正确识别注释行、类型行、字段名行和数据行
- ⚠️ **必须使用两个 `#` 号**：`##namespace` 才会被识别为预处理指令，`#namespace` 会被当作普通注释跳过
- ⚠️ 包含 `##` 指令的行会被跳过，不会生成数据
- ⚠️ 单个 `#` 开头的行仍然是注释，不会被处理

**行号处理示例 1：简单场景**

如果 Sheet 的第一行是 `##namespace` 指令，工具会自动将其跳过并从下一行开始识别：

```
第 0 行: ##namespace Server        ← 跳过此行
第 1 行: ID | 名称 | 类型           ← 识别为注释行
第 2 行: int | string | EnumType   ← 识别为类型行
第 3 行: Id | Name | Type          ← 识别为字段名行
第 4 行: 1001 | 物品A | Type.A     ← 识别为数据行
```

**行号处理示例 2：混合 # 和 ## 的复杂场景**

工具支持在表头行之间插入任意数量的 `#` 注释行，会自动识别真正的数据行：

```
第 0 行: ##namespace Server        ← 预处理指令，添加 using Server;
第 1 行: # 这是注释                ← 单行注释，跳过
第 2 行: ID | 名称 | 职业           ← 识别为注释行（字段说明）
第 3 行: #                         ← 空注释行，跳过
第 4 行: int | string | EnumHeroType ← 识别为类型行
第 5 行: #123                      ← 注释行，跳过
第 6 行: Id | Name | Type          ← 识别为字段名行
第 7 行: 1 | hero1 | HeroType.Knight ← 识别为数据行
第 8 行: #                         ← 注释行，跳过
第 9 行: 2 | hero2 | HeroType.Mage ← 识别为数据行
```

**核心逻辑：**
- `GetDataStartRow()` - 跳过开头的所有 `##` 预处理指令和 `#` 注释，找到第一个实际数据行
- `GetNextNonCommentRow()` - 从指定位置开始查找下一个非注释行（自动跳过 `#` 和 `##`）
- 使用这两个方法动态定位三个头部行：注释行、类型行、字段名行
- 数据起始行 = 字段名行索引 + 1

工具内部使用这套机制确保正确解析表格结构，无论中间插入多少注释行。

#### 1.2 未来扩展

预定义指令系统设计为可扩展的，未来可以添加更多指令：
- `##import` - 导入其他配置文件
- `##define` - 定义常量或宏
- `##region` - 代码区域标记
- 等等...

### 2. 增量编译机制

工具会记录每个 Excel 文件的最后修改时间（存储在 `history.txt`）：

- **首次生成**：处理所有文件
- **后续生成**：只处理修改过的文件
- **强制重新生成**：删除 `history.txt` 文件

**忽略增量检查：**
在 `ignoreHistory.txt` 文件中列出的文件名将总是重新生成（每行一个文件名，不含扩展名）。

### 3. 编译验证

生成完成后，工具会自动编译所有 C# 文件：

- **成功**：显示绿色提示框，保存历史记录
- **失败**：显示红色提示框，输出编译错误详情

**跳过编译验证：**
```bash
dotnet run -- -dir ./Configs -out ./exportcsharp -compile ""
```

### 4. 字符串 Key 的 Dictionary

当第一列是 `string` 类型时，查询方法会使用索引映射：

```csharp
var idx = allItemss.IndexOf(id);
if (idx == -1) return null;
var m = t.GetMethod($"CreateItems_{idx}", ...);
```

### 5. 数据行注释（跳过解析）

**功能说明**：任意数据行的第一列如果以 `#` 或 `##` 开头，该行将被跳过不解析。

- **单个 `#`**：简单注释，直接跳过该行
- **双重 `##`**：预定义指令，先处理指令再跳过该行（参见"预定义指令"章节）

**示例 1：普通数据表**

| ID | 名称 | 价格 |
|---|---|---|
| int | string | int |
| Id | Name | Price |
| 1001 | 新手剑 | 100 |
| #1002 | 钢铁剑（已废弃） | 500 |
| 1003 | 精铁剑 | 800 |

在上述示例中，ID 为 `#1002` 的行将被完全跳过，不会生成到代码中。

**示例 2：参数表（Sheet 名：`#全局参数`）**

| 类型 | 变量名 | 值 | 注释 |
|------|--------|-----|------|
| int | MaxLevel | 100 | 最大等级 |
| #float | OldCritRate | 0.2 | 已废弃的暴击率 |
| string | ServerUrl | https://api.game.com | 服务器地址 |

参数表中以 `#` 开头的行也会被跳过。

**适用场景：**
- ✅ 临时禁用某条配置数据进行测试
- ✅ 保留已废弃的数据作为参考但不生成
- ✅ 在不删除行的情况下快速注释掉数据
- ✅ 策划人员临时屏蔽配置而不影响其他人

**注意事项：**
- 只检查第一列是否以 `#` 或 `##` 开头
- 适用于**所有表格类型**（List 模式、Dictionary 模式、参数表）
- 与 Sheet 名称的 `#` 前缀（参数表）不冲突
- `##` 开头的行会先处理预定义指令（如 `##namespace`）再跳过

### 6. 字段忽略

将字段名或类型设置为 `ignore` 可以跳过该列：

| ID | 废弃字段 | 名称 |
|---|---------|-----|
| int | ignore | string |
| Id | OldField | Name |

### 7. 部分类支持

生成的所有类都使用 `partial` 修饰符，支持在其他文件中扩展：

```csharp
// 自动生成的文件
public partial class Config_Items { }
public partial class Items { }

// 你的扩展文件
public partial class Items
{
    public void CustomMethod() { }
}
```

## 错误处理

### 运行时错误

工具在遇到异常时会显示详细信息：

```
导出失败！
出错的文件：D:\Configs\Items.xlsx
出错的Sheet：Items
System.Exception: Missing cell at ...
```

### 常见错误

| 错误 | 原因 | 解决方案 |
|------|------|----------|
| `Missing cell` | 单元格为空但类型不允许 | 检查 Excel 表格是否有缺失数据 |
| `Unsupported file format` | 不支持的文件格式 | 只支持 `.xls` 和 `.xlsx` |
| 编译错误 | 生成的 C# 代码语法错误 | 检查类型定义是否正确 |

### 调试模式

工具会在控制台输出：
- `Generated <file>` - 开始处理文件
- `Already generated <file>` - 文件未修改，跳过
- `X files cannot be jumped` - 忽略增量检查的文件数量

## 性能优化

### 优化特性

1. **流式读取**：使用 `OptimizedExcelProcessor` 减少内存占用
2. **Tab 缓存**：预先缓存缩进字符串，避免重复创建
3. **延迟加载**：Dictionary 模式使用反射延迟创建对象
4. **增量编译**：跳过未修改的文件

### 大文件处理

对于包含数千行数据的 Excel：
- ✅ 使用 `.xlsx` 格式（支持流式读取）
- ✅ 拆分为多个 Sheet
- ⚠️ 避免在单个 Sheet 中使用过多列（建议 < 50 列）

## 项目集成

### 1. 生成到游戏项目

```bash
dotnet run -- -dir ../Configs -out ../GameProject/src/Shared/Configs -compile ../GameProject/src/Shared/Configs
```

### 2. Git 配置建议

`.gitignore` 文件：
```
# 忽略生成的代码（可选，建议提交）
# exportcsharp/

# 忽略历史文件
history.txt

# 忽略临时文件
~$*.xlsx
```

### 3. CI/CD 集成

在构建脚本中添加：
```bash
cd Tools/Configs/Excel2Code
dotnet run -- -dir ../Configs -out ../../../src/Shared/Configs -compile ../../../src/Shared/Configs
```

## 实用示例

### 示例 1：物品配置表

**Excel（Items.xlsx - Sheet: Items）：**

| 物品ID | 名称 | 类型 | 价格 | 属性加成 |
|-------|------|------|------|---------|
| int | string | int | int | int[] |
| Id | Name | Type | Price | Attributes |
| 1001 | 新手剑 | 1 | 100 | 10,5,0 |
| 1002 | 钢铁剑 | 1 | 500 | 25,12,5 |

**使用代码：**
```csharp
var item = Config_Items.OnGetFrom_dItems(1001);
Console.WriteLine(item.Name);  // 输出：新手剑
Console.WriteLine(item.Attributes[0]);  // 输出：10
```

### 示例 2：技能配置（带向量）

**Excel（Skills.xlsx - Sheet: Skills）：**

| 技能ID | 名称 | 释放位置 | 作用范围 |
|-------|------|---------|---------|
| int | string | v3 | v2[] |
| Id | Name | CastPos | EffectAreas |
| 2001 | 火球术 | 0,1.5,0 | 1,1,2,2 |

**生成代码使用：**
```csharp
var skill = Config_Skills.OnGetFrom_dSkills(2001);
var pos = skill.CastPos;  // DataVector3
var areas = skill.EffectAreas;  // DataVector2[]
```

### 示例 3：枚举类型配置

**Excel（Heroes.xlsx - Sheet: Heroes）：**

| 英雄ID | 英雄名 | 职业类型 | 品质 | 可用技能类型 |
|-------|--------|---------|------|------------|
| int | string | EnumHeroType | EnumItemQuality | EnumSkillType[] |
| Id | Name | Type | Quality | SkillTypes |
| 3001 | 火焰法师 | HeroType.Magic | ItemQuality.Epic | SkillType.Fire,SkillType.AOE |
| 3002 | 剑圣 | HeroType.Physical | ItemQuality.Legendary | SkillType.Melee,SkillType.Single |

**生成代码：**
```csharp
public partial class Heroes
{
    public int Id;
    public string Name;
    public HeroType Type;        // 自动去除了 Enum 前缀
    public ItemQuality Quality;  // 自动去除了 Enum 前缀
    public SkillType[] SkillTypes; // 枚举数组
}

private static Heroes CreateHeroes_3001()
{
    return new Heroes()
    {
        Id = 3001,
        Name = @"火焰法师",
        Type = HeroType.Magic,
        Quality = ItemQuality.Epic,
        SkillTypes = new SkillType[]{SkillType.Fire,SkillType.AOE},
    };
}
```

**使用代码：**
```csharp
var hero = Config_Heroes.OnGetFrom_dHeroes(3001);
Console.WriteLine(hero.Name);  // 输出：火焰法师
Console.WriteLine(hero.Type);  // 输出：HeroType.Magic
Console.WriteLine(hero.Quality); // 输出：ItemQuality.Epic

// 遍历技能类型
foreach (var skillType in hero.SkillTypes)
{
    Console.WriteLine(skillType);
}
```

### 示例 4：全局参数表

**Excel（Configs.xlsx - Sheet: #全局参数）：**

| 类型 | 变量名 | 值 | 注释 |
|------|--------|-----|------|
| int | MaxPlayerLevel | 100 | 玩家最大等级 |
| float | GoldDropRate | 1.5 | 金币掉落倍率 |
| string[] | ServerList | server1,server2 | 服务器列表 |

**使用代码：**
```csharp
Console.WriteLine(Config_Configs.MaxPlayerLevel);  // 100
Console.WriteLine(Config_Configs.GoldDropRate);  // 1.5f
foreach (var server in Config_Configs.ServerList)
{
    Console.WriteLine(server);
}
```

## 注意事项

1. **生成的代码勿手动修改**：文件头部会有 `//本文件为自动生成，请勿手动修改` 警告
2. **Excel 文件命名**：避免使用空格和特殊字符
3. **字段名规范**：必须是有效的 C# 标识符（不能以数字开头）
4. **类型一致性**：同一列的所有数据必须符合声明的类型
5. **编码格式**：Excel 文件建议使用 UTF-8 编码保存
6. **并发安全**：生成的 Dictionary 查询方法不是线程安全的
7. **枚举类型**：
   - 使用 `EnumXXX` 或 `XXXEnum` 格式声明类型
   - 枚举值必须包含完整的类型前缀（如 `HeroType.Magic`）
   - 确保枚举类型已在项目中定义，否则会编译失败
   - 枚举数组中的每个值都需要完整的类型前缀

## 故障排除

### 问题：生成的文件为空
**解决**：检查 Excel 文件是否被其他程序占用（如打开状态）

### 问题：编译失败 - 类型不匹配
**解决**：检查 Excel 中的数据是否符合类型定义（如 int 列是否包含文本）

### 问题：查询返回 null
**解决**：
1. 检查 ID 是否存在于 Excel 中
2. 检查是否调用了正确的查询方法（如 `OnGetFrom_dItems` 而非直接访问 `dItems`）

### 问题：中文乱码
**解决**：确保 Excel 文件以 UTF-8 编码保存

### 问题：枚举类型编译失败
**常见错误**：
```
error CS0246: 未能找到类型或命名空间名"HeroType"
```

**解决方案**：
1. **检查枚举定义**：确保 `HeroType` 枚举已在项目中定义
   ```csharp
   public enum HeroType
   {
       Magic,
       Physical,
       Support
   }
   ```

2. **检查类型命名**：Excel 中使用 `EnumHeroType`，工具会自动转换为 `HeroType`

3. **检查值格式**：Excel 数据行中必须写完整的枚举值
   - ✅ 正确：`HeroType.Magic`
   - ❌ 错误：`Magic`（缺少类型前缀）
   - ❌ 错误：`HeroType.magic`（大小写不匹配）

4. **命名空间问题**：确保生成的配置类能访问到枚举类型（同命名空间或已 using）

## 更新日志

- **v1.0**：初始版本，支持基础类型和集合类型
- **v1.1**：添加增量编译机制
- **v1.2**：优化大文件处理性能（OptimizedExcelProcessor）
- **v1.3**：支持附加表（`=Sheet`）和 partial 类
- **v1.4**：添加数据行注释功能（`#` 开头跳过解析）
- **v1.5**：支持枚举类型（`EnumXXX` / `XXXEnum`）及枚举数组
- **v1.6**：添加预定义指令系统（`##` 前缀），支持 `##namespace` 指令自动添加 using 语句
- **v1.7**：修复混合 `#` 和 `##` 行的解析问题
  - 新增 `GetNextNonCommentRow()` 方法，支持在表头行之间插入任意数量的注释行
  - 修复 `SheetToString` 使用硬编码 `startRow + 3` 导致的行号错误
  - 优化动态行索引计算逻辑，确保正确识别注释行、类型行、字段名行和数据行
  - 修复多sheet中 `##namespace` 指令可能被遗漏的问题

## 相关文件

- `Program.cs` - 命令行入口和参数解析
- `Excel2Code.cs` - 核心生成逻辑
- `OptimizedExcelProcessor.cs` - 优化的 Excel 读取器
- `Utils.cs` - 工具方法（Tab 缓存等）
- `history.txt` - 文件修改历史记录（自动生成）
- `ignoreHistory.txt` - 忽略增量检查的文件列表（手动维护）

## 技术支持

遇到问题请提供以下信息：
1. 完整的错误信息
2. 出错的 Excel 文件结构（截图或示例）
3. 使用的命令行参数
4. 工具版本信息
