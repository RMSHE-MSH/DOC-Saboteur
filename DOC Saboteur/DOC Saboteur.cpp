/**
 * @file DOC Saboteur.cpp
 * @date 15.03.2024
 * @author RMSHE
 *
 * < GasSensorOS >
 * Copyright(C) 2024 RMSHE. All rights reserved.
 *
 * This program is free software : you can redistribute it and /or modify
 * it under the terms of the GNU Affero General Public License as
 * published by the Free Software Foundation, either version 3 of the
 * License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the
 * GNU Affero General Public License for more details.
 *
 * You should have received a copy of the GNU Affero General Public License
 * along with this program.If not, see < https://www.gnu.org/licenses/>.
 *
 * Electronic Mail : asdfghjkl851@outlook.com
 */

#include <iostream>
#include <random>
#include <string>
#include <windows.h>
#include <stdexcept>
#include <pugixml.hpp>
#include <fstream>
#include <sstream>
#include "doc_operate.hpp"

operate op;

// 生成随机字体大小的辅助函数。
int randomFontSize() {
    std::random_device rd;
    std::mt19937 mt(rd());
    std::uniform_int_distribution<int> dist(5, 72); // 范围示例：5pt到72pt
    return dist(mt);
}

// 生成随机字体颜色（十六进制）的辅助函数。
std::string randomFontColor() {
    std::random_device rd;
    std::mt19937 mt(rd());
    std::uniform_int_distribution<int> dist(0, 255);
    char color[7];
    sprintf(color, "%02X%02X%02X", dist(mt), dist(mt), dist(mt));
    return std::string(color);
}

// 生成随机布尔属性（加粗或斜体）的辅助函数。
std::string randomBooleanAttribute() {
    std::random_device rd;
    std::mt19937 mt(rd());
    std::uniform_int_distribution<int> dist(0, 1);
    return dist(mt) ? "true" : "false";
}

// 随机化字体属性函数，用于修改XML内容。
std::string randomizeFontPropertiesEnhanced(const std::string &xmlContent) {
    pugi::xml_document doc;
    if (!doc.load_string(xmlContent.c_str())) return "";

    for (auto paragraph : doc.child("w:document").child("w:body").children("w:p")) {
        for (auto run : paragraph.children("w:r")) {
            auto rPr = run.child("w:rPr");
            if (!rPr) {
                rPr = run.append_child("w:rPr");
            }

            // 设置随机字体大小
            auto sz = rPr.child("w:sz");
            if (!sz) {
                sz = rPr.append_child("w:sz");
                sz.append_attribute("w:val");
            }
            sz.attribute("w:val").set_value(std::to_string(randomFontSize() * 2).c_str());

            // 设置随机字体颜色
            auto color = rPr.child("w:color");
            if (!color) {
                color = rPr.append_child("w:color");
                color.append_attribute("w:val");
            }
            color.attribute("w:val").set_value(randomFontColor().c_str());

            // 随机设置加粗和斜体属性
            auto b = rPr.child("w:b");
            if (!b) {
                b = rPr.append_child("w:b");
                b.append_attribute("w:val");
            }
            b.attribute("w:val").set_value(randomBooleanAttribute().c_str());

            auto i = rPr.child("w:i");
            if (!i) {
                i = rPr.append_child("w:i");
                i.append_attribute("w:val");
            }
            i.attribute("w:val").set_value(randomBooleanAttribute().c_str());
        }
    }

    std::stringstream ss;
    doc.save(ss);
    return ss.str();
}

int main() {
    std::cout << "DOC Saboteur - v0.1;\nPowered by RMSHE - 2024.03.15;\n" << std::endl;

    try {
        std::vector<std::string> file_list = op.find_docx_files("./");
        for (auto iter = file_list.begin(); iter != file_list.end(); ++iter) {
            if (op.copy_file(*iter, "C:/doc_saboteur_backup/" + *iter) == true) {
                std::filesystem::path _docxPath_ = op.renameIfChinese(*iter);

                // 转换为 std::string
                std::string docxPathString = _docxPath_.string();

                // 转换为 const char* 类型
                const char *docxPath = docxPathString.c_str();
                const char *docXmlPath = "word/document.xml";

                std::cout << docxPath << std::endl;

                std::string xmlContent = op.readZipFile(docxPath, docXmlPath);

                // 随机化处理
                std::string modifiedXml = randomizeFontPropertiesEnhanced(xmlContent);
                op.writeZipFile(docxPath, docXmlPath, modifiedXml);

                std::cout << "文档处理完成:" + *iter << std::endl;
            }
        }
    } catch (const std::exception &e) {
        std::cerr << "发生错误: " << e.what() << std::endl;

        system("pause");
        return 1;
    }

    system("pause");
    return 0;
}