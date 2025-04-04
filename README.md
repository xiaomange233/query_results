# 学生成绩分析系统

基于Python开发的成绩分析系统，为教育工作者提供智能化的学生成绩分析解决方案。

[English](README.md) | 简体中文

## 功能特性

- 📊 多维度成绩可视化分析
- 🔍 智能模糊搜索学生档案
- 📈 自动生成成长趋势图表
- 🏫 班级/年级排名对比
- 📁 支持多版本Excel文件解析

## 快速开始
  数据准备
  在项目根目录创建 exams 文件夹
  按规范格式存放Excel考试文件：
  文件扩展名必须为 .xlsx
  第一行为考试场次信息
  第二行为列标题（必须包含"姓名"和"现班"）
  第三行起为学生数据
  
  示例文件结构：
A1: 2023学年第一次月考
A2: 姓名 | 现班 | 语文 | 数学... 
A3: 张三 | 高一(1)班 | 85 | 92...

操作流程
  启动后自动加载exams文件夹数据
  在搜索框输入学生姓名（支持模糊匹配）
  选择正确的学生姓名
  查看左侧成绩详情卡片
  分析右侧可视化趋势图表
  鼠标悬停柱状图查看详细对比

### 环境要求

- Python 3.7+
Windows/macOS/Linux

- 依赖库：
  ```bash
  pip install pandas openpyxl

### 示例文件结构

  └── exams/
    ├── 2023-期中考试.xlsx
    ├── 2023-期末考试.xlsx
    └── 2024-模拟考试.xlsx

### 注意事项

⚠️ ​重要提示
  确保Excel文件格式严格符合要求
  总分列必须命名为"总分"
  各科成绩列建议使用标准学科名称
  文件编码推荐使用UTF-8
  单次加载建议不超过20个考试文件
  学生姓名请使用标准中文姓名
  建议在1366×768及以上分辨率屏幕使用，可获得最佳显示效果。如遇数据显示异常，请检查Excel文件格式是否符合规范。
  
### 开源协议
本项目采用 MIT License，欢迎二次开发和学习使用。如需用于商业场景，请提前联系作者授权。
    
