<div align="center">
  <h1>📦 物料图片匹配系统</h1>
  <p>根据物料清单自动匹配对应的产品图片，支持模糊匹配、多选导出、缩放旋转</p>
  <p>
    <a href="#-功能特性">功能特性</a> •
    <a href="#-下载">下载</a> •
    <a href="#-使用方法">使用方法</a> •
    <a href="#-匹配算法">匹配算法</a> •
    <a href="#-打包为exe">打包</a>
  </p>
  <p>
    <img src="https://github.com/WeiPeng-Pan/Rapid_Imagematch/actions/workflows/build.yml/badge.svg" alt="Build Status">
  </p>
</div>

---

## ✨ 功能特性

| 功能 | 说明 |
|---|---|
| **智能匹配** | 基于注意力机制的模糊匹配算法，自动匹配物料到图片 |
| **多选导出** | 每项可勾选多张图片，导出时全部打包（`_a`, `_b` 后缀） |
| **手动选图** | 匹配结果以缩略图卡片展示，点击即可切换预览 |
| **图片缩放** | 鼠标滚轮缩放，支持 25%~400% |
| **图片旋转** | 90°/180°/270° 任意旋转 |
| **递归扫描** | 自动搜索所有子文件夹中的图片，无需逐层选择 |
| **灵活加载** | 手动选择图片文件夹和 Excel 物料清单 |
| **模板下载** | 内置模板生成，填写即可使用 |
| **结果导出** | 导出为 图片文件夹 + Excel 汇总 + HTML 预览 |
| **列名自适应** | 自动识别物料名称/型号/品牌/参数列（支持中英文别名） |

## 📥 下载

从 GitHub Actions 下载最新打包的 EXE：

👉 [https://github.com/WeiPeng-Pan/Rapid_Imagematch/actions](https://github.com/WeiPeng-Pan/Rapid_Imagematch/actions)

```
1. 点开最新绿色通过的 Build Windows EXE
2. 滑到底部 Artifacts → 点击下载
3. 解压后双击 .exe 运行
```

### 源码运行

```bash
git clone https://github.com/WeiPeng-Pan/Rapid_Imagematch.git
cd Rapid_Imagematch
pip install -r requirements.txt
python3 app_ctk.py
```

## 🚀 使用方法

### 快速上手

```
1. 启动应用 → 点击「选择图片文件夹」选图片目录（自动递归搜索）
2. 点击「加载物料清单」选 Excel 文件
3. 点击「开始匹配」→ 进入主界面
4. 点击「全量匹配」自动匹配所有物料
5. 点击卡片 ✓ 标记勾选多张导出（可选）
6. 点击「导出结果」保存到本地
```

### 多选导出

```
┌──────────┐  ┌──────────┐  ┌──────────┐
│ ✓ 89分   │  │ ○ 72分   │  │ ○ 55分   │
│  (绿色框) │  │  (灰框)   │  │  (灰框)   │
└──────────┘  └──────────┘  └──────────┘
  已勾选导出    未勾选        未勾选

导出结果:
  01-齿轮箱弹性支撑-14_002_003.jpg        (单张)
  02-高速刹车片磨损传感器-490_3711_804_a.jpg  (多张第1)
  02-高速刹车片磨损传感器-490_3711_804_b.jpg  (多张第2)
```

### 支持的 Excel 格式

| 必须列 | 可选列 | 可选列 | 可选列 |
|---|---|---|---|
| **物料名称** | **型号** | **品牌** | **参数** |
| 齿轮箱弹性支撑 | 14/002/003 | ESM | 48763 |
| 高速刹车片磨损传感器 | 490-3711-804 | SVENDBORG | |

> 💡 列名自适应：支持 `物料名称` / `品名` / `NAME`、`型号` / `规格型号` / `MODEL` 等多种写法

### 导出目录结构

```
匹配结果输出/
├── 图片/
│   ├── 01-齿轮箱弹性支撑-14_002_003.jpg
│   ├── 02-高速刹车片磨损传感器-490_3711_804_a.jpg
│   ├── 02-高速刹车片磨损传感器-490_3711_804_b.jpg
│   └── ...
├── 匹配结果汇总.xlsx
└── 预览.html
```

## 🎯 匹配算法

基于注意力机制（Attention Mechanism）的模糊匹配算法：

```
查询串 → 字符串清洗 → jieba分词 + 2-gram → TF-IDF加权
       → 位置衰减 → 完整匹配奖励 → 型号特异性扣分
       → 综合评分 (0-100) → 排序取最优
```

- 中文分词 + 字符 n-gram 双重提取 Token
- TF-IDF 加权：罕见型号编码权重更高
- 位置衰减：开头核心词权重最高
- 型号特异性扣分：防止「分配器 A」匹配到「分配器 B」

## 🔧 打包为 EXE

项目已配置 GitHub Actions 自动打包，推 tag 即触发：

```bash
git tag v1.0
git push origin v1.0
```

手动打包：

```bash
pip install pyinstaller
pyinstaller build.spec
# 输出在 dist/ 目录
```

## 📂 项目结构

```
Rapid_Imagematch/
├── app_ctk.py              # 主程序
├── matching.py              # 核心匹配算法
├── requirements.txt         # 依赖清单
├── build.spec               # 打包配置
├── icon.ico                 # Windows 图标
├── icon.icns                # macOS 图标
├── .github/workflows/
│   └── build.yml            # 自动打包配置
├── .gitignore
└── README.md
```

## 📋 依赖

- Python 3.8+
- openpyxl — Excel 读写
- Pillow — 图片处理
- jieba — 中文分词
- customtkinter — 现代化 GUI

---

<div align="center">
  <p>Rapid_Imagematch · 物料图片自动匹配系统</p>
</div>
