import os
import sys

def open_output_directory(output_dir):
    """打开输出目录"""
    try:
        if sys.platform == 'win32':
            os.startfile(output_dir)
        elif sys.platform == 'darwin':
            os.system(f'open "{output_dir}"')
        else:
            os.system(f'xdg-open "{output_dir}"')
    except Exception as e:
        print(f"⚠️ 无法打开输出目录: {str(e)}")