import pandas as pd
import json
import os
import numpy as np
import sys
import webbrowser
from geopy.geocoders import Nominatim
from tqdm import tqdm
import re
from collections import defaultdict
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
from wordcloud import WordCloud

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

def get_school_location(school_name, cache):
    """获取学校地理位置（使用缓存机制）"""
    # 检查缓存
    if school_name in cache:
        return cache[school_name]
    
    # 使用OpenStreetMap API
    geolocator = Nominatim(user_agent="cengfan_map_app")
    location = None
    city = "未知"
    coords = (0, 0)
    
    # 尝试多种查询格式
    queries = [
        f"{school_name}大学, 中国",
        f"{school_name}, 中国",
        school_name + "大学",
        school_name
    ]
    
    for query in queries:
        try:
            location = geolocator.geocode(query, country_codes='cn', timeout=10)
            if location:
                break
        except Exception as e:
            print(f"⚠️ 查询'{school_name}'时出错: {str(e)}")
            continue
    
    if location:
        # 尝试从地址中提取城市信息
        address = location.address
        # 使用更精确的地址提取方法
        if '市' in address:
            parts = address.split('市')
            city = parts[0] + '市'
            # 尝试获取更精确的位置描述
            if len(parts) > 1 and parts[1].strip():
                address = parts[1].strip().split(',')[0] + ', ' + city
            else:
                address = city
        elif '区' in address:
            parts = address.split('区')
            city = parts[0] + '区'
            address = city
        elif '县' in address:
            parts = address.split('县')
            city = parts[0] + '县'
            address = city
        else:
            # 尝试提取更大的行政区划
            parts = address.split(',')
            if len(parts) > 2:
                city = parts[-3].strip()
                address = city
        
        coords = (location.latitude, location.longitude)
        print(f"✅ 定位成功: {school_name} -> {address}")
    else:
        print(f"⚠️ 无法定位: {school_name}")
        address = "未知位置"
    
    # 保存到缓存
    result = {"city": city, "coords": coords, "address": address}
    cache[school_name] = result
    return result

def generate_html_template(center, markers, output_path):
    """生成优化的HTML地图模板"""
    html = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>同学蹭饭地图</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        body, html {{
            width: 100%;
            height: 100%;
            font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Helvetica Neue", sans-serif;
            overflow: hidden;
        }}
        #map-container {{
            position: relative;
            width: 100%;
            height: 100%;
        }}
        #map {{
            width: 100%;
            height: 100%;
            background: #f5f5f5;
        }}
        .loading-overlay {{
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(255, 255, 255, 0.95);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 1000;
            flex-direction: column;
            text-align: center;
            padding: 20px;
        }}
        .loading-spinner {{
            border: 4px solid #f3f3f3;
            border-top: 4px solid #3498db;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin-bottom: 15px;
        }}
        .map-title {{
            position: absolute;
            top: 10px;
            left: 50%;
            transform: translateX(-50%);
            z-index: 1000;
            background: rgba(255, 255, 255, 0.9);
            padding: 8px 15px;
            border-radius: 20px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            font-size: 16px;
            font-weight: bold;
            backdrop-filter: blur(5px);
        }}
        .retry-btn {{
            margin-top: 20px;
            padding: 8px 16px;
            background: #3498db;
            color: white;
            border: none;
            border-radius: 4px;
            font-size: 14px;
            cursor: pointer;
        }}
        @keyframes spin {{
            0% {{ transform: rotate(0deg); }}
            100% {{ transform: rotate(360deg); }}
        }}
        .leaflet-popup-content {{
            font-size: 14px;
        }}
        .leaflet-popup-content h4 {{
            margin-bottom: 8px;
        }}
    </style>
    <!-- 使用国内CDN加载Leaflet -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.3/leaflet.min.css" />
    <script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.3/leaflet.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet.markercluster/1.5.3/leaflet.markercluster.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet.markercluster/1.5.3/MarkerCluster.min.css" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet.markercluster/1.5.3/MarkerCluster.Default.min.css" />
</head>
<body>
    <div class="map-title">同学蹭饭地图</div>
    <div id="map-container">
        <div id="map"></div>
        <div class="loading-overlay" id="loadingOverlay">
            <div class="loading-spinner"></div>
            <p>正在加载地图...</p>
            <p style="font-size:12px;color:#777;margin-top:10px;">若长时间未加载，请检查网络连接</p>
            <button class="retry-btn" onclick="initMap()">重新加载</button>
        </div>
    </div>

    <script>
        // 地图数据
        var mapData = {{
            center: {json.dumps(center)},
            zoom: 5,
            markers: {json.dumps(markers, ensure_ascii=False)}
        }};

        // 全局地图变量
        var map = null;
        var loadingOverlay = null;

        // 初始化地图
        function initMap() {{
            try {{
                // 确保加载提示可见
                loadingOverlay = document.getElementById('loadingOverlay');
                loadingOverlay.style.display = 'flex';
                
                // 清除旧地图
                if (map) {{
                    map.remove();
                    document.getElementById('map').innerHTML = '';
                }}
                
                // 确保地图容器存在
                if (!document.getElementById('map')) {{
                    throw new Error("地图容器不存在");
                }}
                
                // 创建新地图
                map = L.map('map', {{
                    center: mapData.center,
                    zoom: mapData.zoom,
                    zoomControl: true,
                    preferCanvas: true,
                    tap: false
                }});
                
                // 添加高德地图
                var amap = L.tileLayer('https://wprd0{{s}}.is.autonavi.com/appmaptile?lang=zh_cn&size=1&style=7&x={{x}}&y={{y}}&z={{z}}', {{
                    subdomains: ['1', '2', '3', '4'],
                    attribution: '&copy; 高德地图',
                    maxZoom: 18,
                    minZoom: 3
                }}).addTo(map);
                
                // 添加腾讯地图作为备选
                var qqmap = L.tileLayer('https://rt{{s}}.map.gtimg.com/realtimerender?z={{z}}&x={{x}}&y={{y}}&type=vector&style=0', {{
                    subdomains: ['0', '1', '2'],
                    attribution: '&copy; 腾讯地图',
                    maxZoom: 18,
                    minZoom: 3
                }});
                
                // 创建标记聚类
                var markers = L.markerClusterGroup({{
                    spiderfyOnMaxZoom: false,
                    showCoverageOnHover: false,
                    zoomToBoundsOnClick: true,
                    maxClusterRadius: 60,
                    iconCreateFunction: function(cluster) {{
                        var count = cluster.getChildCount();
                        return L.divIcon({{
                            html: '<div style="background:#1976d2;color:white;border-radius:50%;width:40px;height:40px;display:flex;align-items:center;justify-content:center;font-weight:bold;border:2px solid white">' + count + '</div>',
                            className: 'marker-cluster',
                            iconSize: L.point(40, 40)
                        }});
                    }}
                }});
                
                // 添加标记点
                mapData.markers.forEach(function(loc) {{
                    var marker = L.marker([loc.lat, loc.lng]);
                    var popupContent = `
                        <div style="max-width:300px;padding:10px;">
                            <h4 style="margin:0 0 8px;font-size:15px;">${{loc.address}}</h4>
                            <hr style="margin:6px 0;">
                            <p><b>${{loc.title}}</b>: ${{loc.students.join(', ')}}</p>
                        </div>
                    `;
                    marker.bindPopup(popupContent);
                    markers.addLayer(marker);
                }});
                
                map.addLayer(markers);
                
                // 添加比例尺
                L.control.scale({{imperial: false, metric: true}}).addTo(map);
                
                // 地图加载完成后移除加载提示
                map.whenReady(function() {{
                    setTimeout(function() {{
                        loadingOverlay.style.display = 'none';
                    }}, 500);
                }});
                
                // 添加图层控制
                var baseLayers = {{
                    "高德地图": amap,
                    "腾讯地图": qqmap
                }};
                L.control.layers(baseLayers).addTo(map);
                
            }} catch (error) {{
                console.error("地图初始化失败:", error);
                loadingOverlay.innerHTML = `
                    <h3 style="margin-bottom:15px;color:#e74c3c;">地图加载失败</h3>
                    <p>错误信息: ${{error.message || '未知错误'}}</p>
                    <p style="font-size:12px;color:#777;margin:15px 0;">请检查网络连接后重试</p>
                    <button class="retry-btn" onclick="initMap()">重新加载</button>
                `;
            }}
        }}
        
        // 页面加载完成后初始化地图
        document.addEventListener('DOMContentLoaded', function() {{
            // 确保元素存在
            if (!document.getElementById('map-container')) {{
                document.body.innerHTML = '<div style="padding:20px;text-align:center;"><h3>页面加载错误</h3><p>地图容器未创建</p></div>';
                return;
            }}
            
            // 设置超时检测
            setTimeout(function() {{
                if (!map) {{
                    const loadingOverlay = document.getElementById('loadingOverlay');
                    if (loadingOverlay) {{
                        loadingOverlay.innerHTML = `
                            <h3 style="margin-bottom:15px;color:#e74c3c;">地图加载超时</h3>
                            <p style="font-size:12px;color:#777;margin:15px 0;">可能是网络问题，请尝试重新加载</p>
                            <button class="retry-btn" onclick="initMap()">重新加载</button>
                        `;
                    }}
                }}
            }}, 10000); // 10秒超时
            
            // 初始化地图
            initMap();
        }});
    </script>
</body>
</html>
    """
    
    # 保存HTML文件
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    
    print(f"✅ 已生成优化版蹭饭地图: {output_path}")
    return output_path

def generate_wordcloud(df, output_dir):
    """生成学校词云图"""
    print("正在生成学校分布词云...")
    
    # 解决中文字体问题
    font_path = None
    for font in fm.findSystemFonts():
        if 'SimHei' in font or 'Microsoft YaHei' in font or 'simkai' in font:
            font_path = font
            break
    
    if not font_path:
        print("⚠️ 未找到中文字体，词云可能无法显示中文")
    
    # 生成词云
    school_count = df['学校'].value_counts()
    wordcloud = WordCloud(
        font_path=font_path,
        width=1200,
        height=600,
        background_color='white',
        max_words=50,
        colormap='tab20'
    ).generate_from_frequencies(school_count)
    
    # 保存词云图
    plt.figure(figsize=(15, 8))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis("off")
    plt.title("同学大学分布", fontsize=20, fontproperties={'family': 'SimHei'} if font_path else None)
    plt.tight_layout()
    
    wordcloud_path = os.path.join(output_dir, "学校分布词云.png")
    plt.savefig(wordcloud_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"✅ 词云图已保存: {wordcloud_path}")
    return wordcloud_path

def generate_stats(df, output_dir):
    """生成统计数据表格"""
    print("正在生成统计数据...")
    
    # 城市分布统计
    city_stats = df[df['城市'] != '未知'].groupby('城市').agg({
        '姓名': 'count',
        '学校': lambda x: ', '.join(x.unique())
    }).reset_index()
    city_stats.columns = ['城市', '人数', '学校列表']
    
    # 学校分布统计
    school_stats = df.groupby('学校').agg({
        '姓名': 'count',
        '城市': 'first'
    }).reset_index()
    school_stats.columns = ['学校', '人数', '城市']
    
    # 保存统计结果
    stats_path = os.path.join(output_dir, "大学分布统计.xlsx")
    with pd.ExcelWriter(stats_path) as writer:
        city_stats.to_excel(writer, sheet_name='城市分布', index=False)
        school_stats.to_excel(writer, sheet_name='学校分布', index=False)
        df.to_excel(writer, sheet_name='原始数据', index=False)
    
    print(f"✅ 统计表格已保存: {stats_path}")
    return stats_path

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
    
    # 创建缓存
    cache_path = os.path.join(output_dir, "location_cache.json")
    cache = {}
    if os.path.exists(cache_path):
        try:
            with open(cache_path, "r", encoding='utf-8') as f:
                cache = json.load(f)
            print(f"✅ 已加载缓存 ({len(cache)} 条记录)")
        except:
            pass
    
    # 获取学校地理位置
    print("正在获取学校地理位置...")
    schools = df['学校'].unique()
    locations = {}
    
    for school in tqdm(schools, desc="处理学校"):
        locations[school] = get_school_location(school, cache)
    
    # 保存缓存
    with open(cache_path, "w", encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    print(f"✅ 位置缓存已保存: {cache_path}")
    
    # 添加位置信息到DataFrame
    df['城市'] = df['学校'].map(lambda x: locations[x]['city'])
    df['经纬度'] = df['学校'].map(lambda x: locations[x]['coords'])
    
    # 准备地图数据
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
    
    # 计算有效坐标的中心点
    valid_coords = [loc['coords'] for loc in locations.values() if loc['coords'] != (0, 0)]
    if valid_coords:
        center = np.mean(valid_coords, axis=0).tolist()
    else:
        center = [35.8617, 104.1954]  # 中国中心作为后备
    
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
    if sys.platform == 'win32':
        os.startfile(output_dir)
    elif sys.platform == 'darwin':
        os.system(f'open "{output_dir}"')
    else:
        os.system(f'xdg-open "{output_dir}"')
    
    input("\n按Enter键退出程序...")

if __name__ == "__main__":
    main()