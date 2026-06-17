# 亚马逊SP广告自定义报表生成工具
## 功能介绍
本地离线处理亚马逊 Sponsored Products Campaigns 报表，自动分多Sheet筛选投放问题数据：
1. 广告位无效点击 / 广告位高ACOS
2. 定位广告无效点击 / 定位广告高ACOS / 定位广告0点击低曝光
3. 精准Exact关键词无效点击（点击≥5、ACOS=0）、精准关键词高ACOS
内置固定匹配逻辑：按Campaign ID自动匹配对应ASIN/SKU分行；定位广告ASIN自动生成亚马逊可点击超链接，纯本地运行不上传数据。

## 环境依赖
Python >=3.9
```bash
pip install -r requirements.txt