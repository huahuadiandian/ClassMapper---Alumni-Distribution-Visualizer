import pandas as pd
import os
import sys
import webbrowser
import numpy as np
from tqdm import tqdm

# 导入自定义模块
from config import get_user_config
from geolocation import load_cache, save_cache, get_school_location, add_location_data
from visualization import generate_wordcloud, generate_stats
from html_generator import generate_html_template
from utils import open_output_directory

def main():
    """主程序"""
    # 获取用户配置
    config = get_user_config()
    output_dir = config["output_dir"]
    
    # 读取Excel数据
    try:
        df = pd.read_excel(config["excel_path"])
        if '姓名' not in df.columns or '学校' not in df.columns:
            print("❌ Excel文件中必须包含'姓名'和'学校'两列")
            input("按Enter键退出...")
            return
        print(f"✅ 成功读取Excel数据，共 {len(df)} 条记录")
    except Exception as e:
        print(f"❌ 读取Excel文件失败: {str(e)}")
        input("按Enter键退出...")
        return
    
    # 加载位置缓存
    cache_path = os.path.join(output_dir, "location_cache.json")
    cache = load_cache(cache_path)
    
    # 获取学校地理位置
    print("正在获取学校地理位置...")
    schools = df['学校'].unique()
    locations = {}
    
    for school in tqdm(schools, desc="处理学校"):
        locations[school] = get_school_location(school, cache)
    
    # 保存缓存
    save_cache(cache, cache_path)
    
    # 添加位置信息到DataFrame
    df = add_location_data(df, locations)
    
    # 准备地图数据
    markers = prepare_map_data(df, locations)
    center = calculate_map_center(locations)
    
    # 生成优化的HTML地图
    map_path = os.path.join(output_dir, "蹭饭地图.html")
    generate_html_template(center, markers, map_path)
    
    # 生成词云图
    wordcloud_path = generate_wordcloud(df, output_dir)
    
    # 生成统计表格
    stats_path = generate_stats(df, output_dir)
    
    # 显示结果
    print("\n" + "=" * 50)
    print("🎉 处理完成！生成的文件:")
    print(f"- 蹭饭地图: {map_path}")
    print(f"- 学校词云: {wordcloud_path}")
    print(f"- 统计数据: {stats_path}")
    print("=" * 50)
    
    # 尝试打开地图
    try:
        webbrowser.open(map_path)
        print("已尝试在浏览器中打开蹭饭地图")
    except:
        print("⚠️ 无法自动打开地图，请手动打开HTML文件")
    
    # 打开输出目录
    open_output_directory(output_dir)
    
    input("\n按Enter键退出程序...")

def prepare_map_data(df, locations):
    """准备地图标记数据"""
    markers = []
    for school, group in df.groupby('学校'):
        # 获取位置信息
        loc = locations[school]
        lat, lng = loc['coords']
        address = loc['address']
        
        # 获取该学校的所有学生
        students = group['姓名'].tolist()
        
        markers.append({
            "lat": lat,
            "lng": lng,
            "title": school,
            "students": students,
            "address": address
        })
    return markers

def calculate_map_center(locations):
    """计算地图中心点"""
    valid_coords = [loc['coords'] for loc in locations.values() if loc['coords'] != (0, 0)]
    if valid_coords:
        center = np.mean(valid_coords, axis=0).tolist()
    else:
        center = [35.8617, 104.1954]  # 中国中心作为后备
    return center

if __name__ == "__main__":
    main()