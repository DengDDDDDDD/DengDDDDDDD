# RVR 归零点差异统计（RX）

用于比较两个 Excel 文件中 `RVR` sheet 的 **RX 行**对应的 **归零点**，并按相同信道（如 `5G-48`）计算差值：

- 差值定义：`A表归零点 - B表归零点`
- 输出：
  - `rvr_zero_point_diff.csv`：每个信道的明细
  - `rvr_zero_point_diff_summary.json`：统计摘要（均值/标准差/分位数等），并包含脚本提取过程的 notes，方便定位表格结构差异

## 使用方法

1) 安装依赖：

```bash
pip install -r requirements.txt
```

2) 运行（把两个 Excel 路径换成你的实际文件）：

```bash
python tools/rvr_zero_diff.py "表A.xlsx" "表B.xlsx"
```

如果你的 sheet 名不是 `RVR`，可以指定：

```bash
python tools/rvr_zero_diff.py "表A.xlsx" "表B.xlsx" --sheet "RVR"
```

## 说明

脚本会优先按“结构化表格”解析（有 `信道/Channel` 与 `归零/Zero` 列），如果失败会退回到“启发式解析”：在网格里找包含“归零”的行，并结合附近的 `RX` 标记与信道表头（如 `5G-48`）提取数据。

