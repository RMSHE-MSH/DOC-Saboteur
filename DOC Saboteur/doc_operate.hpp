/**
 * @file doc_operate.hpp
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

#pragma once
#include <zip.h>
#include <pugixml.hpp>
#include <stdexcept>
#include <vector>
#include <filesystem>
#include <locale>
#include <codecvt>
#include <random>

class operate {
public:
    // 从ZIP文件中读取特定文件内容的函数，支持UTF-8编码的路径和文件名。
    std::string readZipFile(const char *archivePath, const char *filePath) {
        int err = 0;
        zip *za = zip_open(archivePath, 0, &err);
        if (!za) {
            throw std::runtime_error("无法打开压缩文件。");
        }

        zip_stat_t sb;
        if (zip_stat(za, filePath, 0, &sb) != 0) {
            zip_close(za);
            throw std::runtime_error("无法获取压缩文件中的文件状态。");
        }

        zip_file *zf = zip_fopen(za, filePath, 0);
        if (!zf) {
            zip_close(za);
            throw std::runtime_error("无法在压缩文件中打开文件。");
        }

        std::vector<char> contents(sb.size);
        zip_fread(zf, contents.data(), sb.size);

        zip_fclose(zf);
        zip_close(za);

        return std::string(contents.begin(), contents.end());
    }

    // 向ZIP文件中写入字符串内容的函数。
    void writeZipFile(const char *archivePath, const char *filePath, const std::string &content) {
        int err = 0;
        zip *za = zip_open(archivePath, ZIP_CREATE, &err);
        if (!za) {
            throw std::runtime_error("无法打开或创建压缩文件。");
        }

        zip_source_t *zs = zip_source_buffer(za, content.c_str(), content.length(), 0);
        if (zip_file_add(za, filePath, zs, ZIP_FL_OVERWRITE) < 0) {
            throw std::runtime_error("无法向压缩文件添加内容。");
        }

        zip_close(za);
    }

    // 删除文件的函数
    bool deleteFile(const std::string &filePath) {
        if (std::remove(filePath.c_str()) == 0) {
            std::cout << "文件已成功删除: " << filePath << std::endl;
            return true;
        } else {
            std::cerr << "无法删除文件: " << filePath << std::endl;
            return false;
        }
    }

    //遍历所有docx文件
    std::vector<std::string> find_docx_files(const std::filesystem::path &start_path) {
        std::vector<std::string> found_files;

        if (!std::filesystem::exists(start_path) || !std::filesystem::is_directory(start_path)) {
            std::cerr << "Path does not exist or is not a directory: " << start_path << std::endl;
            return found_files;
        }

        // 使用try-catch块来捕获和处理遍历时可能抛出的异常
        try {
            for (const auto &entry : std::filesystem::recursive_directory_iterator(start_path)) {
                try {
                    if (entry.is_regular_file()) {
                        if (entry.path().extension() == ".docx") {
                            found_files.push_back(entry.path().string());
                        }
                    }
                } catch (const std::filesystem::filesystem_error &e) {
                    // 如果访问被拒绝，跳过这个文件/目录
                    std::cerr << "Access denied or error accessing: " << entry.path() << ". Skipping...\n";
                }
            }
        } catch (const std::filesystem::filesystem_error &e) {
            // 如果在设置遍历器时遇到问题（例如，起始路径不可访问）
            std::cerr << "Failed to start directory traversal: " << e.what() << std::endl;
        }

        return found_files;
    }

    // 定义copy_file函数
    bool copy_file(const std::filesystem::path &source, const std::filesystem::path &destination) {
        try {
            // 检查源文件是否存在
            if (!std::filesystem::exists(source)) {
                std::cerr << "Source file does not exist: " << source << std::endl;
                return false;
            }

            // 获取目标文件的父路径，并检查是否存在
            auto parent_path = destination.parent_path();
            if (!parent_path.empty() && !std::filesystem::exists(parent_path)) {
                // 如果目标文件的父路径不存在，则创建它
                std::filesystem::create_directories(parent_path);
            }

            // 检查目标文件是否已经存在
            if (std::filesystem::exists(destination)) {
                std::cerr << "Destination file already exists: " << destination << std::endl;
                // 返回false或根据需要修改为覆盖文件
                return false;
            }

            // 执行文件复制操作
            std::filesystem::copy_file(source, destination);
            return true;
        } catch (const std::exception &e) {
            std::cerr << "Error copying file: " << e.what() << std::endl;
            return false;
        }
    }

    // 检查路径中的文件名是否含有中文字符
    bool containsChinese(const std::string &filePath) {
        // 使用C++11的codecvt_utf8来转换字符串为utf-32，便于处理Unicode字符
        std::wstring_convert<std::codecvt_utf8<char32_t>, char32_t> converter;
        std::u32string utf32String;
        try {
            utf32String = converter.from_bytes(filePath);
        } catch (const std::range_error &e) {
            std::cerr << "Exception converting string to UTF-32: " << e.what() << std::endl;
            return false;
        }

        // 遍历每个Unicode字符，检查是否在中文字符的范围内
        for (auto ch : utf32String) {
            if (ch >= 0x4E00 && ch <= 0x9FFF) {
                return true; // 找到中文字符
            }
        }
        return false; // 没有找到中文字符
    }

    // 生成指定长度的随机字符串
    std::string generateRandomString(size_t length) {
        const std::string characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        std::random_device random_device;
        std::mt19937 generator(random_device());
        std::uniform_int_distribution<> distribution(0, characters.size() - 1);

        std::string random_string;
        for (size_t i = 0; i < length; ++i) {
            random_string += characters[distribution(generator)];
        }
        return random_string;
    }

    // 如果文件名含有中文，则重命名该文件，并返回新的文件路径
    std::filesystem::path renameIfChinese(const std::filesystem::path &path) {
        if (containsChinese(path.u8string())) {
            // 获取文件所在目录和扩展名，以便生成新的文件名
            auto parent_path = path.parent_path();
            auto extension = path.extension();

            // 生成新的文件名（不包含扩展名）
            std::string new_name = generateRandomString(10); // 例如，生成10个字符的随机字符串
            auto new_path = parent_path / (new_name + extension.string());

            // 重命名文件
            std::filesystem::rename(path, new_path);
            return new_path; // 返回新的文件路径
        } else {
            return path; // 如果没有重命名，返回原始路径
        }
    }
};