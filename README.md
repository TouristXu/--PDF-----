# 办公自动化工具集

![项目状态](https://img.shields.io/badge/status-active-green.svg)
![Python版本](https://img.shields.io/badge/python-3.10-blue.svg)
![平台](https://img.shields.io/badge/platform-windows-lightgray.svg)

一个基于 Python 开发的办公自动化工具集合，包含 Excel 数据处理、Access 数据库操作和 PDF 编辑等功能。

## 📁 项目结构

```
python办公/
├── Excel高级筛选工具/          # Excel数据筛选工具
│   ├── excel_filter_tool.py    # 主程序源代码
│   ├── excel_filter_tool.exe   # 打包后的可执行文件
│   └── excel_filter_tool.spec  # PyInstaller配置文件
├── XLS转Access工具/            # XLS文件转Access数据库
│   └── xlsx_to_access_fixed.py # 转换工具
├── PDF综合编辑器/              # PDF编辑工具
│   └── PDF编辑器.exe           # PDF编辑器可执行文件
├── 文档/                       # 说明文档
│   ├── Excel高级筛选工具使用说明.md
│   ├── 安装指南.md
│   └── 打包说明.md
└── 依赖安装脚本/               # 环境配置
    ├── install_dependencies.bat
    └── check_env.py
```

## 🛠️ 工具列表

### 1. Excel高级筛选工具

**功能特点：**
- 支持同时处理 1-5 个 Excel 文件
- 支持 `.xls` 和 `.xlsx` 格式
- 智能表头识别（自动检测表头行）
- 合并单元格自动展开处理
- 精确匹配和模糊查找两种模式
- 基于索引的高效搜索
- 结果支持复制到剪贴板
- 表格浏览功能

**使用方式：**
```bash
# 直接运行
./excel_filter_tool.exe

# 或从源代码运行
python excel_filter_tool.py
```

### 2. XLS转Access数据库工具

**功能特点：**
- 批量读取文件夹下所有 `.xls` 文件
- 自动创建数据库表结构
- 支持自定义表头行位置
- 自动处理特殊字符列名

**使用方式：**
```bash
python xlsx_to_access_fixed.py
```

### 3. PDF综合编辑器

**功能特点：**
- PDF 文件编辑
- 页面管理（添加、删除、重排）
- 文本编辑
- 格式转换

**使用方式：**
```bash
./新版PDF综合编辑器/PDF编辑器_部署包/PDF编辑器/PDF编辑器.exe
```

## 📦 安装指南

### 环境要求
- **操作系统**: Windows 10 / 11 (64位)
- **Python版本**: 3.10.x
- **内存**: 至少 2GB RAM
- **磁盘空间**: 至少 500MB 可用空间

### 依赖安装

**方法一：使用批处理脚本（推荐）**
```bash
双击运行 install_dependencies.bat
```

**方法二：手动安装**
```bash
# 升级 pip
python -m pip install --upgrade pip

# 安装核心依赖
python -m pip install pandas openpyxl xlrd pyodbc pywin32

# 如果遇到网络问题，使用国内镜像
python -m pip install pandas openpyxl xlrd pyodbc pywin32 -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### Access驱动安装

如果使用 XLS 转 Access 工具，需要安装 Microsoft Access Database Engine：

1. 运行 `Microsoft_Access_Database_Engine_2019_安装指南.md` 中的安装程序
2. 或从微软官网下载安装：[Microsoft Access Database Engine 2019](https://www.microsoft.com/en-us/download/details.aspx?id=54920)

## 📖 使用说明

### Excel高级筛选工具

1. **启动程序**: 双击 `excel_filter_tool.exe`
2. **添加文件**: 点击"添加Excel文件"按钮，选择最多5个Excel文件
3. **输入关键词**: 在文本框中输入搜索关键词（每行一个）
4. **设置选项**: 勾选"启用模糊查找"进行模糊匹配
5. **开始搜索**: 点击"开始查找"按钮
6. **查看结果**: 在结果窗口中查看匹配数据，支持复制和筛选

### XLS转Access工具

1. 运行命令：`python xlsx_to_access_fixed.py`
2. 输入包含 `.xls` 文件的文件夹路径
3. 输入 Access 数据库文件路径（不存在则自动创建）
4. 程序自动批量处理所有 XLS 文件

### PDF综合编辑器

1. 运行 `PDF编辑器.exe`
2. 通过菜单打开 PDF 文件
3. 使用工具栏进行编辑操作
4. 保存修改后的文件

## 🛡️ 安全性说明

### Windows Defender 提示处理
由于程序未进行数字签名，首次运行时可能会被 Windows Defender 拦截：

1. 点击"更多信息"
2. 点击"仍要运行"
3. 或添加到排除列表：
   - 打开 Windows 安全中心
   - 病毒和威胁防护 → 管理设置
   - 排除项 → 添加或删除排除项
   - 添加相应的 `.exe` 文件

### 第三方杀毒软件
- 将程序添加到信任列表
- 或临时关闭实时防护
- 或提交到杀毒软件厂商进行白名单验证

## ⚡ 性能优化建议

### Excel处理优化
- 单个文件建议不超过 10MB
- 单个文件建议不超过 10 万行数据
- 首次搜索会创建索引（较慢），后续搜索使用缓存（较快）
- 模糊查找会降低搜索速度，建议在必要时启用

### 大文件处理
- 如需处理更大文件，请分批处理
- 关闭其他占用资源的程序
- 确保有足够的内存

## 🐛 常见问题

### Q: 程序无法启动？
A: 
- 确保系统为 Windows 10 或更高版本（64位）
- 关闭杀毒软件重新尝试
- 以管理员身份运行
- 检查是否有足够的磁盘空间

### Q: Excel筛选工具搜索结果为空？
A:
- 检查关键词拼写
- 尝试启用/禁用模糊查找
- 查看文件信息确认数据格式
- 确认 Excel 文件没有被其他程序占用

### Q: XLS转Access工具无法连接数据库？
A:
- 确保已安装 Microsoft Access Database Engine
- 确认数据库文件路径正确
- 检查数据库文件是否被其他程序占用

### Q: 无法复制结果？
A:
- 确保结果窗口已完全加载
- 点击"复制结果"按钮
- 在目标应用中粘贴 (Ctrl+V)
- 如仍失败，重启程序重试

## 📊 技术栈

| 组件 | 版本 | 说明 |
|------|------|------|
| Python | 3.10.x | 编程语言 |
| Tkinter | 内置 | GUI框架 |
| OpenPyXL | latest | XLSX文件处理 |
| xlrd | 1.2.0 | XLS文件处理 |
| pandas | latest | 数据处理 |
| pyodbc | latest | Access数据库连接 |
| pywin32 | latest | Excel COM接口 |
| PyInstaller | 6.19.0 | 打包工具 |

## 🔧 开发与打包

### 重新打包 Excel筛选工具

**方法一：使用批处理脚本**
```bash
双击运行 build_exe.bat
```

**方法二：手动打包**
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --distpath . excel_filter_tool.py
```

### 打包参数说明
- `--onefile`: 打包成单个 exe 文件
- `--windowed`: 不显示控制台窗口
- `--distpath .`: 输出到当前目录

## 📝 版本历史

| 版本 | 日期 | 更新内容 |
|------|------|----------|
| v1.0 | 2026-04-20 | 初始版本，包含Excel筛选工具和XLS转Access工具 |

## 🚀 未来计划

- [ ] 添加 Excel 数据导出功能
- [ ] 支持更多文件格式（CSV、JSON等）
- [ ] 添加正则表达式搜索
- [ ] 优化大文件处理性能
- [ ] 添加数据统计和分析功能
- [ ] 增加批量处理任务队列

## 📞 联系与支持

如有问题或建议，请：
1. 查看项目中的说明文档
2. 检查 README 中的故障排除部分
3. 确认系统要求是否满足
4. 尝试重新打包或重新下载

## 📄 许可证

本项目仅供内部使用。

---

**最后更新**: 2026-06-09  
**版本**: v1.0  
**状态**: ✅ 可正常使用
