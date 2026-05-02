# 龙之谷互刷巢穴集结规划工具

[English](README.md) | [中文](README.zh-CN.md)

本工具用于根据 Excel 输入自动计算每周队伍刷图计划。

它会读取门票与周委托数据，使用 MILP 求解，并输出一个按执行顺序排好的计划，目标包括：

1. 每场战斗消耗队长 1 张票。
2. 优先最大化全队总周委托完成数。
3. 在总数相同前提下，尽量提高队内最低完成数。
4. 排序时考虑换图成本，并尽量避免整队重组（全队换角色）。

## 目标列表

`狮蝎, 海龙, K博士, 主教, 巨人, 守卫, 火山, 迷雾, 卡伊伦, 格拉诺, 代达罗斯, 台风金, 双生`

## 动态成员

成员由输入目录中的文件名自动识别，按成对文件匹配：

- `<member>_票.xlsx`
- `<member>_委托.xlsx`

代码中不再写死成员名，可直接分享给其他队伍使用。

当只有 3 名成员时：

- 非 `双生` 战斗自动按 3 人编队
- `双生` 仍为 2 人

## 输入格式

### `<member>_票.xlsx`（sheet: `Characters`）

| 列名 | 含义 |
| --- | --- |
| `角色` | 角色名（文件内唯一） |
| `<target>` | 该目标门票数量（整数） |
| `选择` | 选择袋子(期限制) |

### `<member>_委托.xlsx`（sheet: `Quests`）

| 列名 | 含义 |
| --- | --- |
| `角色` | 需与 `<member>_票.xlsx` 中角色一致 |
| `<target>` | `1` 表示该角色该目标有周委托，否则 `0` |

数值单元格允许留空，留空会按 `0` 处理。

## 使用方式

```powershell
# 安装依赖
pip install -r requirements.txt

# 首次运行：若 input/ 不存在或为空，会自动从 templates/ 复制模板
python run.py

# 或直接运行主程序
python -m src.main --input-dir input --out schedule.xlsx
```

## 输出文件

`schedule.xlsx` 包含以下工作表：

- `Schedule`
  - 战斗执行顺序
  - 角色切换单元格橙色高亮
  - 本场完成周委托的角色名加粗
  - 目标列在每次切换地图组时交替换色（浅灰 / 浅粉）
- `Summary`
  - 每位成员完成数、参战数、角色种类数、切换次数
- `Legend`
  - 颜色与符号说明

## 模型与排序说明

- 门票约束严格生效（专用票与通用票都不会透支）。
- 周委托按 `(成员, 角色, 目标)` 每周最多记一次。
- 目标函数为字典序：先总完成数，再平衡性。
- 后处理会移除无贡献战斗并重新排序。
- 排序成本包含：
  - 角色切换成本
  - 换图成本
  - 整队重组成本站（高惩罚）

## 测试与校验

```powershell
# 生成可复现实验数据
python -m tests.generate_test_case

# 校验排程输出是否合法
python -m tests.validate_schedule --input-dir tests/input --schedule tests/output/schedule.xlsx

# 新增特性回归测试（动态成员、空值按0、目标双色）
python -m tests.test_dynamic_features
```

## 项目结构

```text
DN_Tools/
├── README.md
├── README.zh-CN.md
├── requirements.txt
├── run.py
├── run.bat
├── generate_templates.py
├── src/
│   ├── config.py
│   ├── optimize.py
│   ├── schedule.py
│   ├── templates.py
│   └── main.py
└── tests/
    ├── generate_test_case.py
    ├── validate_schedule.py
    ├── test_dynamic_features.py
    ├── count_reteams.py
    └── show_reteams.py
```
