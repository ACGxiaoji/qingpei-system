# 青赔资料管理系统

铁塔租地与青苗补偿场景的资料归档与协议生成工具，Python + Tkinter + SQLite + 离线 OCR，Windows 免安装版。

> 本仓库为脱敏版：项目名、地名已替换为占位符，不含任何真实人员信息与业务数据。

## 功能
- 读取人员资料、证件照、补偿 Excel，自动归档
- 离线 OCR 批量生成 Word 协议与 Excel 补偿表
- SQLite 资料库：登记、检索、汇总导出
- 局域网多机访问（内置 HTTP 服务，无需额外依赖）

## 效果
- 原来手工一天的资料整理，现在十几分钟完成

## 技术点
- 离线 OCR（rapidocr_onnxruntime），不依赖外网
- 批量任务队列 + 异常重试
- Word / Excel 模板化生成（python-docx / openpyxl）
- 免安装打包（PyInstaller），双击即用

## 运行
```bash
pip install -r requirements.txt
python start_qingpei.py
```

## 文件
- `青赔系统.py`：主程序（GUI）
- `agreement_to_table.py`：协议解析与补偿表生成
- `qingpei_db.py`：SQLite 数据层
- `青赔局域网管理系统.py`：局域网 Web 管理端
- `start_qingpei.py`：启动入口
