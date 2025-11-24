import argparse
import json
import os
import sys

import pandas as pd

# 默认配置
DEFAULT_JSON = "zh.json"
DEFAULT_EXCEL = "translation_work.xlsx"


def get_file_paths(args_json, args_excel):
    """
    获取文件路径：优先使用命令行参数，如果没有，则尝试使用交互式输入
    """
    # 1. 确定 JSON 路径
    if args_json:
        json_path = args_json
    else:
        # 交互式询问
        user_input = input(f"请输入 JSON 文件名 (默认: {DEFAULT_JSON}): ").strip()
        json_path = user_input if user_input else DEFAULT_JSON

    # 2. 确定 Excel 路径
    if args_excel:
        excel_path = args_excel
    else:
        # 交互式询问
        user_input = input(f"请输入 Excel 文件名 (默认: {DEFAULT_EXCEL}): ").strip()
        excel_path = user_input if user_input else DEFAULT_EXCEL

    return json_path, excel_path


def json_to_excel(json_file, excel_file):
    print(f"📖 读取 JSON: {json_file}")

    if not os.path.exists(json_file):
        print(f"❌ 错误: 找不到文件 {json_file}")
        return

    try:
        with open(json_file, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"❌ JSON 读取失败: {e}")
        return

    rows = []
    for file_path, items in data.items():
        for original, translation in items.items():
            rows.append(
                {
                    "文件路径 (勿改)": file_path,
                    "原文": original,
                    "译文": translation,
                    "状态": "已翻译" if translation else "待翻译",
                }
            )

    df = pd.DataFrame(rows)
    try:
        df.to_excel(excel_file, index=False, engine="openpyxl")
        print(f"✅ 转换成功！已生成 Excel: {excel_file}")
        print(f"📊 总计条目: {len(df)}")
    except Exception as e:
        print(f"❌ Excel 保存失败: {e}")


def excel_to_json(excel_file, json_file):
    print(f"📖 读取 Excel: {excel_file}")

    if not os.path.exists(excel_file):
        print(f"❌ 错误: 找不到文件 {excel_file}")
        return

    try:
        df = pd.read_excel(excel_file, engine="openpyxl", dtype=str)
        df.fillna("", inplace=True)
    except Exception as e:
        print(f"❌ Excel 读取失败: {e}")
        return

    json_data = {}
    count = 0
    for _, row in df.iterrows():
        file_path = row.get("文件路径 (勿改)")
        original = row.get("原文")
        translation = row.get("译文")

        # 简单校验
        if not file_path or not original:
            continue

        if file_path not in json_data:
            json_data[file_path] = {}

        json_data[file_path][original] = translation
        count += 1

    try:
        with open(json_file, "w", encoding="utf-8") as f:
            json.dump(json_data, f, ensure_ascii=False, indent=4)
        print(f"✅ 转换成功！已更新 JSON: {json_file}")
        print(f"📊 处理条目: {count}")
    except Exception as e:
        print(f"❌ JSON 保存失败: {e}")


def main():
    # 配置命令行参数解析
    parser = argparse.ArgumentParser(description="JSON <-> Excel 互转工具")
    parser.add_argument("-j", "--json", help="指定 JSON 文件路径")
    parser.add_argument("-e", "--excel", help="指定 Excel 文件路径")
    parser.add_argument(
        "mode",
        nargs="?",
        choices=["to_excel", "to_json"],
        help="直接指定模式: to_excel 或 to_json",
    )

    args = parser.parse_args()

    # 如果命令行指定了模式，直接运行
    if args.mode == "to_excel":
        j, e = get_file_paths(args.json, args.excel)
        json_to_excel(j, e)
        return
    elif args.mode == "to_json":
        j, e = get_file_paths(args.json, args.excel)
        excel_to_json(e, j)
        return

    # 否则进入交互模式
    print("--- 汉化文件转换器 ---")
    print("1. JSON 转 Excel (去翻译)")
    print("2. Excel 转 JSON (回填)")

    choice = input("请选择 (1/2): ").strip()

    if choice == "1":
        # 这里传入 None，让函数内部去询问用户文件名
        j, e = get_file_paths(args.json, args.excel)
        json_to_excel(j, e)
    elif choice == "2":
        j, e = get_file_paths(args.json, args.excel)
        excel_to_json(e, j)
    else:
        print("❌ 无效输入")


if __name__ == "__main__":
    main()
