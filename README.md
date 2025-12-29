# MdDollarMathToWord

一个 **Microsoft Word VBA 宏**，用于将 Word 文档中的 **Markdown `$` 数学公式**
（`$...$`、`$$...$$`）一键转换为 **Word 原生公式**。

本项目仅包含 **一个宏代码文件（.bas）**，无插件、无依赖，复制即可使用。

---

## 功能
- 支持行内公式：`$a^2 + b^2 = c^2$`
- 支持块级公式：`$$\int_0^1 x^2 dx$$`
- 转换为 Word 原生公式（可双击编辑）
- 支持转换选区或全文

> ⚠️ 当前版本 **只支持 Markdown 的 `$` 语法**，不支持 `\(...\)`、`\[...\]` 等其它 LaTeX 写法。

---

## 使用方法

1. 打开 Microsoft Word
2. `Alt + F11` 打开 VBA 编辑器
3. `File → Import File...`，导入 `MdDollarMathToWord.bas`
4. 回到 Word
5. 选中文本（可选），运行宏：
