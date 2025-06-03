# ConfigExcel

## 简介 (Introduction)

ConfigExcel 是一款基于 C# 代码热更新方案（如 ILRuntime / HybridCLR）设计的工具。它能够自动将 Excel 表格数据转换为 C# 类代码并直接填充数据，从而避免了运行时的序列化和反序列化开销。生成的代码为标准 C# 代码，可无缝应用于 PC、安卓、iOS、微信小游戏等多种平台。

## 核心优势 (Core Advantages)

*   **零GC、零IO开销**：生成的代码在访问数据时无需进行文件读取或复杂的反序列化操作。
*   **懒加载**：数据在首次访问时才加载，优化启动性能和内存占用。
*   **易于上手**：无需学习新的数据格式或配置方式，直接使用 Excel 进行数据管理。
*   **跨平台兼容**：生成的 C# 代码具有良好的跨平台特性。
*   **热更新友好**：生成的C#代码可以直接集成到热更新的DLL中，方便数据和代码的同步更新。

## 主要功能 (Key Features)

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

*   **Sheet 名为 "类名 变量名" 格式**:
    *   当 Sheet 名的格式为“单词 空格 单词”（例如："PlayerData mPlayerData"）时，Sheet 名会被解析为“类名”和“变量名”。
    *   **成员变量**: 当表格除去表头（注释、变量名、数据类型行）后只有一行有效数据时，该 Sheet 会被识别为类的单个成员变量。
    *   **类字典**: 如果表格中有多行数据，该 Sheet 会被识别为该类的字典，键的类型由第一列决定，值为类实例，即 `Dictionary<第一列的数据类型, 类实例>`。
    *   **类列表**: 如果表格第一行第一列中包含 `list` 关键字，该 Sheet 会被识别为该类的列表，即 `List<类实例>`。

*   **Sheet 名为单个词 (类名)**:
    *   当 Sheet 名只有一个单词时（例如："GlobalConfig"），该单词会被识别为类名。
    *   这种情况下，生成的变量名会根据具体内容自动识别并添加前缀：
        *   `m类名` (例如：`mGlobalConfig`)，通常用于单个实例对象。
        *   `d类名` (例如：`dGlobalConfig`)，通常用于字典。
        *   `l类名` (例如：`lGlobalConfig`)，通常用于列表。

### 支持的数据类型
*   支持各种 C# 基础数据类型 (如 `int`, `string`, `bool`, `float` 等)，需要确保 Excel 中填充的内容与声明的类型匹配。
*   支持基础类型的数组 (如 `int[]`, `string[]`)，同样需要确保填充内容合法。
*   支持 Key 和 Value 均为基础类型的字典 (如 `Dictionary<int, string>`)，需要确保填充内容合法。

## 环境要求 (Prerequisites/Requirements)

请确保已安装 .NET 5.0 或更高版本。

## 使用方法 (Usage)

您可以通过以下几种方式运行 `Excel2Code.exe`：

1.  **导出指定目录下的所有 Excel 文件：**
    打开命令行工具，执行以下命令（假设 `Excels` 是存放 Excel 文件的目录）：
    ```bash
    Excel2Code.exe -dir Excels
    ```

2.  **导出指定的若干个 Excel 文件：**
    打开命令行工具，执行以下命令，列出所有需要导出的文件路径：
    ```bash
    Excel2Code.exe Excels/test.xlsx Excels/other1.xlsx Excels/other2.xlsx
    ```

3.  **通过交互模式导出：**
    直接双击 `Excel2Code.exe` 文件，程序会提示您输入包含 Excel 文件的目录路径。

## 示例 (Examples)

在 `Excel2Code.exe` 工具的执行目录下（通常在 `bin/Debug/net5.0/` 或类似路径下），您可以找到：
*   示例 Excel 文件位于该执行目录下的 `Excels/` 子目录内。这些示例覆盖了目前支持的各种配置写法。
*   生成的 C# 代码样例位于该执行目录下的 `Codes/` 子目录内。
建议参考这些示例以深入了解具体用法和配置规则。

## 许可证 (License)

本项目采用 MIT 许可证。

## 贡献 (Contributing)

欢迎参与贡献！您可以通过提交 Pull Request 或创建 Issue 的方式参与到项目中来。
