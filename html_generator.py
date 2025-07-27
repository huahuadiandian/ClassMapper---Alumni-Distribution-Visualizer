import json
import os

def generate_html_template(center, markers, output_path):
    """生成优化的HTML地图模板（保持原样不变）"""
    # 这里保持原始HTML字符串完全不变
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