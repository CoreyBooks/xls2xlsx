# 发布指南

## 本地验证

```bash
# 1. 语法检查
flake8 src/ tests/

# 2. 运行测试
pytest -v

# 3. 手动运行验证
python main.py test_data/

# 4. 打包测试
pyinstaller --onefile --name xls2xlsx --distpath ./dist main.py
./dist/xls2xlsx.exe
```

## GitHub 自动发布

```bash
git add .
git commit -m "release: v1.0.0"
git tag v1.0.0
git push origin main --tags
```

GitHub Actions 自动：
1. 运行 pytest 测试
2. 执行 flake8 检查
3. PyInstaller 构建 `xls2xlsx.exe`
4. 合并源码构建独立 `xls2xlsx.py` 脚本
5. 创建 GitHub Release 并上传 exe + py 双文件

## Release 文件说明

每次 Release 会包含两个文件：

| 文件 | 适用人群 | 说明 |
|------|----------|------|
| `xls2xlsx.exe` | **最终用户** | 单文件，无需 Python，双击运行 |
| `xls2xlsx.py` | **Python 用户** | 独立单文件脚本，`pip install pywin32` 后 `python xls2xlsx.py` 即可运行 |

## 用户下载后使用

### exe 用户
1. 下载 Release 中的 `xls2xlsx.exe`
2. 放到需要转换的文件夹中
3. 双击运行
4. 按提示完成扫描、确认、转换

### py 脚本用户
1. 下载 Release 中的 `xls2xlsx.py`
2. 确保已安装 Python >= 3.9 和 Microsoft Excel
3. 运行 `pip install pywin32>=306`
4. 运行 `python xls2xlsx.py`（详见 README.md）
