"""工具函数：文件格式检测、路径处理。"""

import os

# 支持的格式映射：后缀 → 格式标识
ALLOWED_EXTENSIONS = {
    ".docx": "docx",
    ".doc": "docx",
    ".pdf": "pdf",
    ".md": "md",
}


def is_valid_file(filepath):
    """检查文件是否为支持的格式。"""
    ext = os.path.splitext(filepath)[1].lower()
    return ext in ALLOWED_EXTENSIONS


def detect_format(filepath):
    """检测文件格式，返回格式标识（docx/pdf/md），不支持返回 None。"""
    ext = os.path.splitext(filepath)[1].lower()
    return ALLOWED_EXTENSIONS.get(ext)


def get_output_path(input_path, target_ext, custom_dir=None):
    """根据输入路径和目标扩展名计算输出路径。

    target_ext: "pdf" / "docx" / "md"（不含点）
    custom_dir: 如果指定，输出到该目录；否则与源文件同目录。
    """
    base = os.path.splitext(os.path.basename(input_path))[0]
    out_name = base + "." + target_ext
    if custom_dir:
        return os.path.join(custom_dir, out_name)
    return os.path.join(os.path.dirname(input_path), out_name)
