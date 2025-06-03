# ConfigExcel

## 简介 (Introduction)

ConfigExcel 是一款基于 C# 代码热更新方案（如 ILRuntime / HybridCLR）设计的工具。它能够自动将 Excel 表格数据转换为 C# 类代码并直接填充数据，从而避免了运行时的序列化和反序列化开销。生成的代码为标准 C# 代码，可无缝应用于 PC、安卓、iOS、微信小游戏等多种平台。

## 核心优势 (Core Advantages)

*   **零GC、零IO开销**：生成的代码在访问数据时无需进行文件读取或复杂的反序列化操作。
*   **懒加载**：数据在首次访问时才加载，优化启动性能和内存占用。
*   **易于上手**：无需学习新的数据格式或配置方式，直接使用 Excel 进行数据管理。
*   **跨平台兼容**：生成的 C# 代码具有良好的跨平台特性。
*   **热更新友好**：生成的C#代码可以直接集成到热更新的DLL中，方便数据和代码的同步更新。

## 推荐使用方法：在 Unity 编辑器中使用 (Recommended Usage: Using in Unity Editor)

对于在 Unity 项目中使用此工具的开发者，我们推荐直接在编辑器内集成和使用，这样可以更便捷地管理和生成代码。

此方法允许您直接在 Unity 编辑器中从 Excel 文件生成 C# 代码，无需使用独立的 `Excel2Code.exe` 程序。

### 设置步骤

1.  将 `UnityDemo/Assets/Editor/AutoExcel/` 目录下的以下文件复制到您 Unity 项目的 `Assets/Editor/` 文件夹下的任意子目录中：
    *   `ConfigExcel.cs`
    *   `Excel2Code.cs`
    *   `Utils.cs`
2.  将 `UnityDemo/Assets/Editor/AutoExcel/EPPlus.dll` 文件复制到您的 Unity 项目中 (例如，放到 `Assets/Plugins/` 或 `Assets/Editor/` 下的某个目录)。这是在 Unity 环境下解析 Excel 文件所需的库。

### Unity 中的使用方法

完成上述设置后，Unity 编辑器顶部菜单栏会出现 "Tools > ConfigExcel" 选项。点击该菜单项会打开一个窗口，您可以在其中指定包含 Excel 文件的输入目录、C# 代码的输出目录以及需要忽略的文件列表（逗号分隔，无需扩展名）。点击 "导出Excel" 按钮即可开始生成代码。

### 与独立版 (.exe) 的主要区别

*   **Excel 解析库 (Excel Parsing Library):** Unity 版本使用 `EPPlus.dll`，而独立版使用 NPOI。因此，`EPPlus.dll` 是 Unity 集成所必需的。
*   **无代码编译 (No Code Compilation):** 此集成脚本不包含编译生成后 C# 代码的功能。您需要依赖 Unity自身的编译流程。
*   **无历史记录/增量生成 (No History/Incremental Generation):** Unity 版本不使用 `history.txt`，每次都会重新生成所有指定文件。
*   **Sheet 命名规则差异 (Sheet Naming Rule Differences):**
    *   不支持 `|` 作为 Sheet 名称中的分隔符 (仅支持空格)。
    *   不支持以 `=` 开头的 Sheet 用于扩展数据。
    *   Unity 版本目前不会显式忽略名为 `Sheet1`, `Sheet2` 等的默认工作表，但如果它们不符合其他解析规则，也可能不会被处理。
*   **类型替换差异 (`drepalace`):** Unity 版本中可用的类型替换较少 (例如，不支持 `dic`, `v2[]`, `v3[]`, `luatable`, `luacode` 的直接替换)。
*   **自定义数据类处理 (Custom Data Class Handling):** 如果您在 Excel 中定义了需要工具实例化的自定义类 (例如，类型列填写为 `MyData`)，该类在 Unity 版本中需要有一个 `public void FromString(string excelData)` 方法来从 Excel单元格的字符串内容初始化自身。例如：`public class MyData { public void FromString(string excelData) { /* 解析 excelData 并填充字段 */ } }`。
*   **命名空间 (Namespace):** Unity 版本生成的 C# 代码不包含外层命名空间 (例如 `namespace ConfigExcel {...}` 被移除)。
*   **输出信息 (Output Messages):** 使用 `UnityEngine.Debug.Log` 输出日志。

## 主要功能 (Key Features)

**注意：** 以下描述的许多功能和规则主要基于独立版 `Excel2Code.exe`。Unity 编辑器集成版本在某些方面（如Sheet命名、部分类型支持、无编译步骤等）存在差异。具体差异请参考“推荐使用方法：在 Unity 编辑器中使用”部分下的“与独立版 (.exe) 的主要区别”表格。核心的数据转换逻辑是相似的。

### 导出 Excel 文件
*   **导出目录中的所有 Excel 文件**:
    ```bash
    Excel2Code.exe -dir Excels
    ```
*   **导出指定的 Excel 文件**:
    ```bash
    Excel2Code.exe Excels/test.xlsx Excels/test2.xlsx
    ```

### Sheet 表命名规则与C#代码生成

*   **以 `#` 开头的 Sheet (基础类型集合)**:
    *   Sheet 名称以 `#` 开头的表格，会被识别为基础类型集合。
    *   表格结构：第一列为数据类型，第二列为变量名，第三列为具体内容。
    *   当内容为 `int`/`string`/`bool` 类型时，第一列（数据类型）可以为空，工具会自动推断类型。

*   **不以 `#` 开头的 Sheet (数据表)**:
    *   Sheet 名称不以 `#` 开头的表格，会被识别为数据表。
    *   表格结构：第一行为注释，第二行为变量名，第三行为数据类型。

*   **Sheet 名为 "类名 变量名" 或 "模块名 类名 变量名" 格式**:
    *   当 Sheet 名的格式为两段式“NameOne NameTwo”（或使用 `|` 分隔：“NameOne | NameTwo”）时，`NameTwo` 会被作为类名，`NameOne` 通常代表模块或分类，解析后的变量名为 `mNameTwo` (单个实例) / `dNameTwo` (字典) / `lNameTwo` (列表)。
    *   当 Sheet 名的格式为三段式“NameOne NameTwo NameThree”（或使用 `|` 分隔：“NameOne | NameTwo | NameThree”）时，`NameThreeData` 会被作为类名，`NameOne` 和 `NameTwo` 通常代表模块或分类，解析后的变量名为 `mNameThreeData` / `dNameThreeData` / `lNameThreeData`。
    *   **成员变量**: 当表格除去表头（注释、变量名、数据类型行）后只有一行有效数据时，该 Sheet 会被识别为类的单个成员变量。
    *   **类字典**: 如果表格中有多行数据，该 Sheet 会被识别为该类的字典，键的类型由第一列决定，值为类实例，即 `Dictionary<第一列的数据类型, 类实例>`。
    *   **类列表**: 如果表格第一行（注释行）第一列中包含 `list` 关键字，该 Sheet 会被识别为该类的列表，即 `List<类实例>`。

*   **Sheet 名为单个词 (类名)**:
    *   当 Sheet 名只有一个单词时（例如："GlobalConfig"），该单词会被识别为类名。
    *   这种情况下，生成的变量名会根据具体内容自动识别并添加前缀：
        *   `d类名` (例如：`dGlobalConfig`)，通常用于字典 (`Dictionary<key, GlobalConfig>`)。
        *   `m类名` (例如：`mGlobalConfig`)，当代码生成为列表时 (例如 `List<GlobalConfig>`)，或作为单个实例对象时（后者通常源于如“成员变量”描述中所述的具有单行数据的Sheet），变量名均使用 `m` 前缀。

*   **以 `=` 开头的 Sheet (扩展数据表)**:
    *   这类 Sheet 用于扩展其前方最近一个主数据表（定义了字典的 Sheet）的数据。它们的内容会合并到主数据表的字典中。例如，如果 `MyDataSheet` 定义了一个字典，并且其后紧跟着 `=SheetPart2` 和 `=SheetPart3`，那么 `=SheetPart2` 和 `=SheetPart3` 中的数据将被添加到 `MyDataSheet` 的字典里。

*   **以 `Sheet` 开头的 Sheet (忽略处理)**:
    *   以 Excel 默认命名模式 `Sheet` 开头（例如 `Sheet1`, `Sheet2` 等）的 Sheet 将被忽略，不进行处理。

### 支持的数据类型
*   支持各种 C# 基础数据类型 (如 `int`, `string`, `bool`, `float` 等)，需要确保 Excel 中填充的内容与声明的类型匹配。
*   支持基础类型的数组 (如 `int[]`, `string[]`)，同样需要确保填充内容合法。
*   支持 Key 和 Value 均为基础类型的字典 (如 `Dictionary<int, string>`)，需要确保填充内容合法。
*   支持用户自定义的复杂类型，如果其名称以 `Data` 开头或结尾（例如 `MyCustomData`），或者使用特定的短语如 `v2`, `v3`, `v2[]`, `v3[]`（这些会分别映射到 `DataVector2`, `DataVector3`, `DataVector2[]`, `DataVector3[]`）。工具会尝试将单元格内容构造成 `new ClassName()` 或 `new ClassName[]{}` 的形式。用户需要确保这些自定义类在项目中已定义。
*   支持类型别名：在 Excel 中可使用 `map` 或 `dic` 来代表 `Dictionary` 类型。同时，`luatable` 和 `luacode` 也会被直接转换为 `string` 类型。例如，类型列中填写 `map<int,string>` 等同于 `Dictionary<int,string>`。

## 环境要求 (Prerequisites/Requirements)

### 对于独立版 `Excel2Code.exe`
*   请确保已安装 .NET 5.0 或更高版本。

### 对于 Unity 编辑器集成
*   需要将 `EPPlus.dll` 添加到项目中（`UnityDemo` 文件夹内已提供此文件）。
*   对 Unity 编辑器本身的版本无特殊要求，能够正常运行您的 Unity 项目即可。

## 备选方法：使用独立版 Excel2Code.exe (Alternative Method: Using Standalone Excel2Code.exe)

除了在 Unity 编辑器中直接使用外，您也可以使用独立的 `Excel2Code.exe` 程序来生成代码。此独立版本拥有一些特有功能，例如通过 `history.txt` 实现的**增量生成**（自动跳过未更改的 Excel 文件以提高效率）、将生成的代码**直接编译成 DLL**，以及更全面的**Sheet命名规则**和**类型别名**支持。当您需要在 Unity 环境之外处理 Excel 文件，或者需要这些特定功能时，可以选用此方法。

您可以通过以下几种方式运行 `Excel2Code.exe`：

1.  **导出指定目录下的所有 Excel 文件：**
    打开命令行工具，执行以下命令（以 `Excels` 作为存放 Excel 文件的目录为例）：
    ```bash
    Excel2Code.exe -dir Excels -out ./GeneratedCode -compile ./GeneratedCode -ignore SecretTable,OldData
    ```
    *   `-dir <Excel目录>`: 必需参数，指定包含 Excel 文件的目录。
    *   `-out <输出目录>`: 可选参数，指定生成的 C# 代码文件 (`.cs`) 的输出目录。如果未提供此参数，则不会生成代码文件。
    *   `-compile <编译目录>`: 可选参数，指定包含要编译的 C# 代码文件的目录（通常与 `-out` 目录相同）。如果提供此参数，工具会在代码生成后尝试将该目录中的所有 `.cs` 文件编译成一个临时的 `temp.dll`。编译成功会提示，编译失败会显示错误。这个过程有助于快速验证生成的代码是否语法正确，并且为需要动态加载DLL的热更新方案提供了一种可能性。如果编译失败，会输出详细错误信息。当编译成功后，`temp.dll` 会被自动删除，它主要用于即时验证。如果未提供此参数，则不执行编译步骤。
    *   `-ignore <文件名列表>`: 可选参数，提供一个逗号分隔的 Excel 文件名列表（无需扩展名），这些文件将被忽略处理。例如：`-ignore SecretTable,OldData`。

2.  **导出指定的若干个 Excel 文件：**
    打开命令行工具，执行以下命令，列出所有需要导出的文件路径。这种方式下，如果未指定 `-out` 参数，生成的 C# 代码将默认输出到程序运行目录下的 `./exportcsharp` 文件夹中。
    ```bash
    Excel2Code.exe Excels/test.xlsx Excels/other1.xlsx Excels/other2.xlsx -out ./SpecificCode
    ```
    同样可以配合使用 `-compile` 参数 (如果 `-out` 未指定，`-compile` 将作用于 `./exportcsharp` 目录)。`-ignore` 参数仅在与 `-dir` 参数一同使用时生效。

3.  **通过交互模式导出：**
    直接双击 `Excel2Code.exe` 文件，程序会提示您输入包含 Excel 文件的目录路径。

## 示例 (Examples)

在 `Excel2Code.exe` 工具的执行目录下（通常在 `bin/Debug/net5.0/` 或类似路径下），您可以找到：
*   示例 Excel 文件位于该执行目录下的 `Excels/` 子目录内。这些示例覆盖了目前支持的各种配置写法。
*   生成的 C# 代码样例位于该执行目录下的 `Codes/` 子目录内。
建议参考这些示例以深入了解具体用法和配置规则。这些示例主要演示了独立版 `Excel2Code.exe` 的配置方法。对于 Unity 版本，Excel 表格本身的结构和数据填写方式基本一致，但部分高级命名规则或类型支持可能存在差异，请务必参考“与独立版 (.exe) 的主要区别”部分的说明。

## 许可证 (License)

本项目采用 MIT 许可证。

## 贡献 (Contributing)

欢迎参与贡献！您可以通过提交 Pull Request 或创建 Issue 的方式参与到项目中来。
