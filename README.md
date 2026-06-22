<div align="center">

# 📊 Coupon Summary Tool

**代金券汇总工具 · 按日汇总多地区商品与代金券数据**

<br>

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)
![OpenPyXL](https://img.shields.io/badge/OpenPyXL-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)

<br>

自动解析多地区销售数据 · 按日聚合汇总 · 生成带颜色分类的 Excel 报表

[功能特性](#功能特性) · [快速启动](#快速启动) · [输入输出](#输入输出) · [项目结构](#项目结构)

</div>

---

## 功能特性

| 功能 | 说明 |
|------|------|
| 📅 按日汇总 | 自动按日期聚合商品和代金券数据 |
| 🏪 多地区支持 | 兰考、民权两个地区独立汇总 |
| 🎨 颜色分类 | Excel 输出按类型着色，一目了然 |
| 🖱️ GUI 界面 | 图形化操作，无需命令行 |
| 📦 单文件打包 | 支持 PyInstaller 打包为 exe |

---

## 快速启动

```bash
# 安装依赖
pip install pandas openpyxl

# 运行
python 代金券汇总工具.py

# 打包 exe
pip install pyinstaller
pyinstaller --onefile --windowed --name "代金券汇总工具" 代金券汇总工具.py
```

---

## 输入输出

### 输入格式

Excel 文件，需包含商品名称、代金券金额、地区、日期等字段。

### 输出结果

按日期分 sheet，每个 sheet 包含该日各地区的商品和代金券汇总，不同类型用颜色区分。

---

## 项目结构

```
coupon-summary-tool/
├── 代金券汇总工具.py       # 主程序（GUI + 汇总逻辑）
└── 代金券汇总工具.exe       # 打包后的可执行文件
```

---

## 开源许可

MIT © [H1nk5](https://github.com/H1nk5)
