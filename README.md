## 运行说明

1. 确保考勤表已准备好（默认使用弹窗选择，或通过 `-i/--input` 指定路径）。
2. 安装依赖：
   ```bash
   python3 -m pip install -r requirements.txt
   ```
3. 执行脚本（默认弹出文件选择窗口，并在脚本/可执行所在目录的 `outputs/` 下生成 `原名-加班统计.xlsx`）：
   ```bash
   python3 main.py
   ```
   Excel 结果与运行日志都会存放在 `outputs/` 目录，可用于排查问题。

### 可选参数

- `-i, --input`：直接指定排班表路径，例如 `python3 main.py -i 数据/考勤.xlsx`
- `-o, --output`：指定输出路径（相对路径会写到同级 `outputs/` 目录），例如 `python3 main.py -o 自定义名称.xlsx`
- `--no-dialog`：禁用选择文件弹窗，常用于纯命令行环境；此时若未指定 `-i` 会回退查找 `10月考勤.xlsx`

脚本会把新增的“n号加班”列附到表格末尾，并在单元格备注中写明班次、扣减的休息时间与最终工时。若排班格式存在问题，原排班单元格会以红色背景标记并加上错误说明。*** End Patch
