"""
批量清理 doc/docx 文件名前缀的工具

命令行模式:
    rename.exe -d D:\docs                      # 指定待处理目录
    rename.exe -k D:\my_keywords.txt           # 指定关键词文件
    rename.exe --dry-run                       # 预览模式，不实际重命名

双击运行:
    直接双击 exe，进入交互模式，按提示操作即可

规则:
    - 只处理 .doc 和 .docx 文件
    - 从 keywords.txt 读取关键词（一行一个）
    - 只删除出现在文件名【开头】的关键词，精确匹配
    - 如果多个关键词都匹配开头，取最长的那个（贪婪匹配，避免误删）
"""

import io
import os
import sys
import argparse


def get_exe_dir() -> str:
    """获取 exe 所在目录（兼容 PyInstaller 打包和直接运行）"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load_keywords(filepath: str) -> list[str]:
    """读取关键词文件，返回按长度降序排列的关键词列表（长的优先匹配）"""
    if not os.path.isfile(filepath):
        print(f"[错误] 关键词文件不存在: {filepath}")
        return []

    keywords = []
    with open(filepath, "r", encoding="utf-8-sig") as f:
        for line in f:
            word = line.strip()
            # 跳过空行和注释行
            if word and not word.startswith("#"):
                keywords.append(word)

    if not keywords:
        print("[警告] 关键词文件为空，没有需要删除的词汇")
        return []

    # 按长度降序排列，确保优先匹配最长的关键词
    keywords.sort(key=len, reverse=True)
    return keywords


def find_matching_prefix(filename_stem: str, keywords: list[str]) -> str | None:
    """查找文件名开头匹配的关键词，返回匹配到的最长关键词"""
    for kw in keywords:
        if filename_stem.startswith(kw):
            return kw
    return None


def process_directory(target_dir: str, keywords: list[str], dry_run: bool = False) -> dict:
    """扫描目录，处理所有 doc/docx 文件，返回统计结果"""
    result = {"renamed": 0, "skipped": 0, "errors": 0}
    changed_list = []    # 改动详情: (原文件名, 新文件名, 匹配关键词)
    unchanged_list = []  # 未改动文件名列表
    error_list = []      # 失败/冲突详情: (文件名, 原因)

    if not os.path.isdir(target_dir):
        print(f"[错误] 目录不存在: {target_dir}")
        return result

    valid_extensions = {".doc", ".docx"}

    files = [f for f in os.listdir(target_dir) if os.path.isfile(os.path.join(target_dir, f))]
    doc_files = [f for f in files if os.path.splitext(f)[1].lower() in valid_extensions]

    if not doc_files:
        print(f"[提示] 目录中没有 doc/docx 文件: {target_dir}")
        return result

    print(f"找到 {len(doc_files)} 个 doc/docx 文件")
    print(f"加载了 {len(keywords)} 个关键词")
    print("-" * 60)

    for filename in sorted(doc_files):
        stem, ext = os.path.splitext(filename)
        matched_kw = find_matching_prefix(stem, keywords)

        if matched_kw is None:
            result["skipped"] += 1
            unchanged_list.append(filename)
            continue

        new_stem = stem[len(matched_kw):].lstrip()

        # 清理后文件名为空的情况
        if not new_stem.strip():
            reason = "删除关键词后文件名为空"
            print(f"[跳过] {reason}: {filename}")
            result["skipped"] += 1
            unchanged_list.append(filename)
            continue

        new_filename = new_stem + ext
        old_path = os.path.join(target_dir, filename)
        new_path = os.path.join(target_dir, new_filename)

        # 检查目标文件是否已存在
        if os.path.exists(new_path):
            reason = f"目标文件已存在: {new_filename}"
            print(f"[冲突] {reason}，跳过: {filename}")
            result["errors"] += 1
            error_list.append((filename, reason))
            continue

        if dry_run:
            print(f"[预览] {filename}  ->  {new_filename}  (删除: \"{matched_kw}\")")
            result["renamed"] += 1
            changed_list.append((filename, new_filename, matched_kw))
        else:
            try:
                os.rename(old_path, new_path)
                print(f"[完成] {filename}  ->  {new_filename}  (删除: \"{matched_kw}\")")
                result["renamed"] += 1
                changed_list.append((filename, new_filename, matched_kw))
            except OSError as e:
                reason = str(e)
                print(f"[失败] {filename}: {reason}")
                result["errors"] += 1
                error_list.append((filename, reason))

    # 输出汇总报告
    print()
    print("=" * 60)
    mode_label = "预览" if dry_run else "执行"
    print(f"  {mode_label}结果汇总")
    print("=" * 60)
    print(f"  doc/docx 文件总数: {len(doc_files)}")
    renamed_label = "待重命名" if dry_run else "已重命名"
    print(f"  {renamed_label}:    {result['renamed']}")
    print(f"  未改动:      {result['skipped']}")
    print(f"  失败/冲突:   {result['errors']}")

    if changed_list:
        print()
        print(f"--- {'待重命名' if dry_run else '已重命名'}详情 ({len(changed_list)} 个) ---")
        for i, (old, new, kw) in enumerate(changed_list, 1):
            print(f"  {i:3d}. {old}")
            print(f"    -> {new}")
            print(f"       删除关键词: \"{kw}\"")

    if error_list:
        print()
        print(f"--- 失败/冲突详情 ({len(error_list)} 个) ---")
        for i, (name, reason) in enumerate(error_list, 1):
            print(f"  {i:3d}. {name}")
            print(f"       原因: {reason}")

    if unchanged_list:
        print()
        print(f"--- 未改动文件 ({len(unchanged_list)} 个) ---")
        for i, name in enumerate(unchanged_list, 1):
            print(f"  {i:3d}. {name}")

    print()
    return result


def run_cli():
    """命令行模式"""
    exe_dir = get_exe_dir()
    default_kw = os.path.join(exe_dir, "keywords.txt")

    parser = argparse.ArgumentParser(description="批量清理 doc/docx 文件名前缀")
    parser.add_argument("-d", "--dir", default=".", help="待处理的文件目录 (默认: 当前目录)")
    parser.add_argument("-k", "--keywords", default=default_kw, help=f"关键词文件路径 (默认: {default_kw})")
    parser.add_argument("--dry-run", action="store_true", help="预览模式，不实际修改文件")
    args = parser.parse_args()

    target_dir = os.path.abspath(args.dir)
    keywords_file = os.path.abspath(args.keywords)

    print(f"目标目录: {target_dir}")
    print(f"关键词文件: {keywords_file}")
    print(f"模式: {'预览 (dry-run)' if args.dry_run else '实际执行'}")
    print("=" * 60)

    keywords = load_keywords(keywords_file)
    if not keywords:
        return
    process_directory(target_dir, keywords, dry_run=args.dry_run)


def run_interactive():
    """交互模式（双击 exe 时进入）"""
    exe_dir = get_exe_dir()

    print("=" * 60)
    print("  批量清理 doc/docx 文件名前缀工具")
    print("=" * 60)
    print()

    # 1. 选择目标目录
    default_dir = exe_dir
    user_input = input(f"请输入文件目录路径 (直接回车使用当前目录 {default_dir}): ").strip()
    target_dir = user_input if user_input else default_dir
    target_dir = os.path.abspath(target_dir)

    if not os.path.isdir(target_dir):
        print(f"[错误] 目录不存在: {target_dir}")
        input("\n按回车键退出...")
        return

    # 2. 选择关键词文件
    default_kw = os.path.join(exe_dir, "keywords.txt")
    user_input = input(f"请输入关键词文件路径 (直接回车使用 {default_kw}): ").strip()
    keywords_file = user_input if user_input else default_kw
    keywords_file = os.path.abspath(keywords_file)

    keywords = load_keywords(keywords_file)
    if not keywords:
        input("\n按回车键退出...")
        return

    print()
    print(f"目标目录: {target_dir}")
    print(f"关键词文件: {keywords_file}")
    print(f"关键词数量: {len(keywords)}")
    print()

    # 3. 先预览
    print(">>> 预览模式：以下文件将被重命名")
    print("=" * 60)
    preview = process_directory(target_dir, keywords, dry_run=True)

    if preview["renamed"] == 0:
        print("\n没有需要重命名的文件。")
        input("\n按回车键退出...")
        return

    # 4. 确认执行
    print()
    confirm = input("确认执行重命名？(输入 y 确认，其他取消): ").strip().lower()
    if confirm != "y":
        print("已取消操作。")
        input("\n按回车键退出...")
        return

    print()
    print(">>> 正式执行重命名")
    print("=" * 60)
    process_directory(target_dir, keywords, dry_run=False)

    input("\n按回车键退出...")


def setup_console_encoding():
    """设置 Windows 控制台编码为 UTF-8，避免中文乱码"""
    if sys.platform == "win32":
        os.system("chcp 65001 >nul 2>&1")
        # 确保 stdout/stderr 使用 UTF-8
        if hasattr(sys.stdout, 'buffer'):
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


def main():
    setup_console_encoding()
    # 有命令行参数时走CLI模式，没有参数时走交互模式
    if len(sys.argv) > 1:
        run_cli()
    else:
        run_interactive()


if __name__ == "__main__":
    main()
