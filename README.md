# Word题库转换工具（Word Question Bank Converter）

一个强大高效的自动化批量转换工具，用于将大量非结构化的 Word 题库文档（.docx）自动转换为结构化的 Excel 文件，方便导入数据库、在线考试系统或小程序。

如果这个工具帮到了您，请给它一个 ⭐️！

## ✨ 主要特性

- **批量处理**：一键处理整个文件夹内所有 Word 文档，高效省时。
- **题型全面**：完美支持单选题、多选题和判断题的识别与解析。
- **灵活定制**：可通过修改配置或脚本，适配不同的 Word 文档格式和题型标识。
- **高精度解析**：基于 Pandoc 的文件转换和关键选项模糊查找逻辑，准确提取题目、选项和答案。
- **多种输出**：生成整洁的 XLSX 文件，可直接作为数据库导入表使用。
- **智能清理**：内置强大的文本预处理算法，自动标准化格式，去除冗余符号和错乱内容。
- **自动化流程**：自动关闭相关进程，防止资源冲突，支持错误日志记录，便于维护。

## 🚀 快速开始

### 1. 安装依赖

**Windows 系统**
- Microsoft Office
- [Pandoc](https://pandoc.org/)
- curl（用于获取网络时间）

### 2. 准备 Word 文件

- 将所有需要转换的 Word 文档（.docx）放入一个文件夹，例如 `./word_files/`
- 确保文档格式相对统一（如答案以"答案："开头）

### 3. 运行转换工具

1. 下载并解压工具包
2. 运行主程序（AutoHotkey 脚本）
3. 按提示选择输入文件夹和输出 Excel 文件路径
4. 等待转换完成

### 4. 获取结果

程序运行结束后，将在指定路径生成 Excel 文件，使用 Excel 打开即可看到结构化的题库数据。

## 📊 输出格式

转换后的 XLSX 文件包含以下列：

| 列名 | 说明 | 示例 |
|------|------|------|
| id | 题目唯一ID（自动生成） | 1 |
| type | 题型 (single/multiple/tf) | single |
| question | 题目正文 | 什么是Python？ |
| option_a_b_c_d | A选项内容 | 一种编程语言 |
| answer | 正确答案 | A 或 ABC |

## 🛠 文本预处理逻辑

- 多轮字符串递归替换（StrReplaceAll），彻底清理嵌套、连续、错乱的分隔符和特殊符号。
- 自动标准化题型标识、选项格式、答案字段。
- 支持自定义清洗规则，适应复杂题库文本。

## ❓ 常见问题

**Q: 运行需要什么环境？**  
A: 需要 Windows 系统并安装 Microsoft Office。

**Q: 支持哪些题型？**  
A: 完美支持单选题、多选题和判断题。

**Q: 转换精度如何？**  
A: 基于 Pandoc 和高精度正则表达式，转换准确率很高。

**Q: 能处理复杂的 Word 格式吗？**  
A: 支持大多数常见格式，特殊格式可通过自定义规则微调。

## 📥 下载地址

**最新版本下载：**  
[百度网盘下载链接](https://pan.baidu.com/s/16-3bwhj75IoLFWXyuzWwgA?pwd=d2k5)  
提取码: d2k5

## 📺 视频教程

[B站视频教程](https://www.bilibili.com/video/BV1phpwznEuo/)

## 🤝 如何贡献

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

## 📄 许可证

本项目采用 MIT 许可证 - 查看 LICENSE 文件了解详情。

## 💡 致谢

感谢 [Pandoc](https://pandoc.org/) 库提供了强大的 Word 文档转换能力。

## 🛠 开发者

邮箱: 2055770551@qq.com

如需进一步定制或有任何问题，欢迎通过 Issues 或邮箱联系。