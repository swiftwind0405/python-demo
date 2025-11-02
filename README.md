## 运行说明

1. 确保本目录下存在 `10月考勤.xlsx`（或通过 `-i/--input` 指定其它路径）。
2. 安装依赖：
   ```bash
   python3 -m pip install -r requirements.txt
   ```
3. 执行脚本（默认在同目录生成 `10月考勤-加班统计.xlsx`）：
   ```bash
   python3 main.py
   ```

### 可选参数

- `-i, --input`：输入排班表路径，例如 `python3 main.py -i 数据/考勤.xlsx`
- `-o, --output`：指定输出路径，例如 `python3 main.py -o 输出/考勤-加班统计.xlsx`

脚本会把新增的“n号加班”列附到表格末尾，并在单元格备注中写明班次、扣减的休息时间与最终工时。若排班格式存在问题，原排班单元格会以红色背景标记并加上错误说明。*** End Patch
