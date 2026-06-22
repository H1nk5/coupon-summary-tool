<div align="center">

# 📊 Coupon Summary Tool

**代金券汇总工具 · 按日汇总多地区商品与代金券数据**

<br>

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)

<br>

按日汇总兰考 / 民权的商品和代金券数据，生成带颜色分类的 Excel 报表

[功能特性](#功能特性) · [快速启动](#快速启动) · [使用说明](#使用说明)

</div>

---

## 功能特性

- 📅 按日期自动汇总商品数据和代金券数据
- 🏪 支持兰考、民权两个地区
- 🎨 Excel 输出带颜色分类，一目了然
- 🖱️ GUI 界面，操作简单

---

## 快速启动

### 安装依赖

```bash
pip install pandas openpyxl
```

### 运行

```bash
python main.py
```

### 打包 exe

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "代金券汇总工具" main.py
```

---

## 使用说明

1. 运行程序，选择输入的 Excel 文件
2. 选择输出目录
3. 程序自动按日期和地区汇总
4. 输出带颜色分类的 Excel 报表

---

## 开源许可

MIT © [H1nk5](https://github.com/H1nk5)
