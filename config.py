import os

def get_user_config():
    """获取用户配置信息"""
    config = {
        "excel_path": "",
        "output_dir": "蹭饭地图结果"
    }
    
    print("=" * 50)
    print("同学蹭饭地图生成器 v1.5")
    print("=" * 50)
    
    # 获取Excel文件路径
    while not config["excel_path"]:
        path = input("请输入Excel文件路径（直接拖拽文件到此处或输入路径）：").strip('"')
        if os.path.isfile(path) and path.endswith(('.xlsx', '.xls')):
            config["excel_path"] = path
        else:
            print("❌ 文件不存在或不是Excel文件，请重新输入")
    
    # 创建输出目录
    os.makedirs(config["output_dir"], exist_ok=True)
    
    return config