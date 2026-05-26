# Well Extractor — 地下水统测信息提取系统

AI 视觉 + OCR 双引擎地下水井点调查数据提取，从微信小程序长截图自动提取字段并导出 Excel。

## 功能

- **双引擎识别**：EasyOCR（本地离线） + MiMo 视觉（云端 AI）
- **赛博朋克 UI**：PySide6 霓虹风格桌面应用
- **坐标匹配**：野外编号模糊匹配坐标文件，经纬度/XY/高程来自坐标文件
- **模板导出**：按统测表模板输出 Excel，自动水位标高计算

## 引擎选择

| 引擎 | 方式 | 速度 | 精度 | 成本 |
|------|------|------|------|------|
| EasyOCR | 本地离线 | 慢（加载10s+） | 字段级 | 0 |
| MiMo 视觉 | 云端 API | 快（~3s/张） | 表单级 | ~1000 tokens/张 |

## 项目结构

```
well-extractor/
├── desktop_app_v2.py    # PySide6 GUI 主程序
├── ocr_engine.py        # OCR 引擎（EasyOCR + MiMo 视觉 API）
├── keygen_gui.py        # 许可证密钥生成器
├── license.py           # 机器码许可证系统
└── gen_word.py          # Word 文档生成
```

## 运行

```bash
# 安装依赖
pip install PySide6 easyocr openpyxl pillow numpy

# 配置 MiMo（可选）
# 编辑 mimo_config.json 填入 Xiaomi MiMo API key

# 启动
python desktop_app_v2.py
```

## 数据文件路径

```
D:\地下水\信息截图\           # 微信小程序截图
D:\地下水\坐标.xlsx           # 坐标文件
D:\地下水\大同朔州忻州_地下水统测表.xlsx  # Excel 模板
```

## License

Proprietary — 内部使用
