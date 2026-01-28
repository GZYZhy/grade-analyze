# grade-analyze
一款成绩分析工具（Web 方式，跨平台运行）。

## 运行环境
- Python 3.9
- macOS / Windows

## 安装依赖
在项目根目录执行：
```
pip install -r requirements.txt
```

## 启动应用
```
streamlit run app.py
```

## 功能概览
- 导入输入成绩单 / 通常成绩单
- 成绩趋势展示（PPT 风格循环）
- 导出分析报告（单文件或按学生拆分）
- 学生对比
- 统计分析（阈值、进退步、百分位线、导出 Excel/图片）
- 数据管理（修正、删除、导出全量数据）
- 多班级切换
- Highlight 科目
- 软件打开密码保护
