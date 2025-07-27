import json
import os
from geopy.geocoders import Nominatim

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

def load_cache(cache_path):
    """加载位置缓存"""
    cache = {}
    if os.path.exists(cache_path):
        try:
            with open(cache_path, "r", encoding='utf-8') as f:
                cache = json.load(f)
            print(f"✅ 已加载缓存 ({len(cache)} 条记录)")
        except:  # noqa: E722
            pass
    return cache

def save_cache(cache, cache_path):
    """保存位置缓存"""
    with open(cache_path, "w", encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    print(f"✅ 位置缓存已保存: {cache_path}")

def add_location_data(df, locations):
    """添加位置信息到DataFrame"""
    df['城市'] = df['学校'].map(lambda x: locations[x]['city'])
    df['经纬度'] = df['学校'].map(lambda x: locations[x]['coords'])
    return df