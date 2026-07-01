# 液体火箭质量、质心和转动惯量计算工具

本项目是一个基于 **QJ 1080A-1997《液体火箭质量、质心和转动惯量计算方法》** 的液体火箭质量特性计算工具。

项目采用“Excel 输入/输出 + Python 计算内核”的技术路线：Excel 只负责工程参数录入和结果交付，复杂计算全部在 Python 中实现，便于检查、测试、维护和版本管理。

## 功能

- 从 `input_template.xlsx` 读取结构化输入参数。
- 使用平行轴定理合成部组件质量、质心和转动惯量。
- 按级间/助推器分离时间计算不变质量突变。
- 计算可抛质量随时间的分段变化。
- 实现 QJ 1080A 中贮箱第一类/第二类半径模型。
- 对推进剂体积、质心和转动惯量随液位变化进行数值积分。
- 计算增压气体质量特性。
- 计算贮箱可变质量随飞行时间变化。
- 按 QJ 标准使用启动后剩余量 `Mef` 作为发动机工作段消耗起点。
- 输出从箱底算起的贮箱液位结果和运输状态部段合成结果。
- 输出带格式的 Excel 结果报表和 PDF 曲线图。
- 输入模板包含边框、下拉框、批注、数据校验和预留空白输入行。
- 输出报表包含结果总览、关键时间点汇总和关键事件前后对比。
- 使用 `pytest` 对核心计算函数进行测试。

## 文件结构

```text
rocket_mass_properties.py   核心计算库
run_calculation.py          命令行运行入口
create_input_template.py    生成/刷新 Excel 输入模板
create_starship_example.py  生成 Starship/Super Heavy 公开数据近似样例
input_template.xlsx         可编辑的输入模板
input_starship_example.xlsx Starship/Super Heavy 近似输入样例
```

计算结果会生成到 `output/` 目录，该目录已加入 `.gitignore`，不会提交到仓库。

## 环境要求

推荐使用 Python 3.10 或更新版本。

依赖包：

```text
numpy
pandas
openpyxl
matplotlib
pytest
```

安装依赖：

```powershell
pip install numpy pandas openpyxl matplotlib pytest
```

## 使用方法

生成或刷新输入模板：

```powershell
python create_input_template.py
```

编辑 `input_template.xlsx` 后运行计算：

```powershell
python run_calculation.py
```

生成 Starship/Super Heavy 公开数据近似样例：

```powershell
python create_starship_example.py
```

运行 Starship/Super Heavy 近似样例：

```powershell
python run_calculation.py --input input_starship_example.xlsx --output output/starship_mass_properties_result.xlsx --plots output/starship_mass_properties_plots.pdf
```

默认输出文件：

```text
output/mass_properties_result.xlsx
output/mass_properties_plots.pdf
```

指定输入和输出路径：

```powershell
python run_calculation.py --input input_template.xlsx --output output/mass_properties_result.xlsx --plots output/mass_properties_plots.pdf
```

跳过 PDF 绘图：

```powershell
python run_calculation.py --no-plots
```

## 输入表格说明

`input_template.xlsx` 包含以下工作表：

- `总体参数`：火箭总级数、助推器数量、飞行时间范围、时间步长、液位积分点数。
- `级事件`：各子级/助推器的点火时间、关机时间和分离时间。
- `不变质量`：结构、设备、有效载荷等不变质量部组件的质量、质心和中心主惯量。
- `可抛质量`：整流罩等可抛部件的质量、质心、惯量和抛离时间。
- `贮箱基本参数`：贮箱密度、秒流量、加注量、启动后剩余量、关机剩余量、坐标原点偏移和满箱参考量。
- `贮箱几何_第一类`：QJ 1080A 第一类贮箱半径模型参数。
- `贮箱几何_第二类`：QJ 1080A 第二类内部扣除体积半径模型参数。
- `外行导管`：外行导管对滚转惯量 `Jx` 的贡献。
- `运输状态`：运输状态部段组件输入表。

输入模板的格式化特性：

- 每个输入工作表预留到第 `500` 行，空白行也带边框、底色、下拉框和基础数据校验。
- 需要选择的字段使用下拉框，例如 `类型`、`几何类型`、`是否启用扣除`。
- 关键表头带批注说明，鼠标悬停可查看填写提示。
- `填写说明` 工作表集中说明坐标系、单位、填写顺序、贮箱ID规则和级间分离规则。
- 输入区使用浅黄色底色和蓝色字体；备注列使用浅灰底色。
- 首行冻结并开启筛选，列宽已按字段类型放宽。

角度字段，例如 `beta11_deg`、`beta12_deg`，统一使用“度”为单位。程序内部会自动转换为弧度。

`贮箱基本参数` 中的 `几何类型` 当前支持：

- `QJ第一类`：使用 `贮箱几何_第一类` 和可选 `贮箱几何_第二类` 中的 QJ 1080A 分段半径模型。
- `等效圆柱`：只使用 `等效半径R` 和 `等效高度H` 进行圆柱贮箱积分，不依赖 `贮箱几何_第一类`。

`等效圆柱` 适合公开资料不足、只需要做工程近似验证的型号样例，例如 Starship/Super Heavy。

## 输出表格说明

`mass_properties_result.xlsx` 包含以下工作表：

- `结果总览`
- `关键时间点汇总`
- `关键事件前后对比`
- `输入检查`
- `不变质量结果`
- `贮箱液位结果`
- `贮箱箱底液位结果`
- `运输状态结果`
- `不变质量时间序列`
- `可变质量时间序列`
- `可抛质量时间序列`
- `全箭时间序列`

输出工作簿的格式化特性：

- 输出表格使用与输入模板一致的深蓝表头、边框、冻结首行和自动筛选。
- 结果区使用浅绿色底色，和输入表的浅黄色输入区区分。
- 输出表自动加宽列宽，并对质量、坐标、时间、转动惯量等字段设置数值格式。
- 输出表在数据区后额外预留空白带框线区域，便于临时查看和扩展。

`结果总览` 包含：

- 输入文件
- 时间范围和时间点数量
- 贮箱数量和级事件数量
- 初始全箭质量
- 初始质心 `X/Y/Z`
- 初始转动惯量 `Jx/Jy/Jz`
- 最大质量、最小非零质量和最终质量
- 输入检查状态

`关键时间点汇总` 会自动汇总：

- 飞行开始和飞行结束
- 各子级/助推器点火、关机和分离时刻
- 可抛质量抛离时刻
- 对应时刻的全箭 `M, X, Y, Z, Jx, Jy, Jz`

`关键事件前后对比` 会自动输出事件前后变化量：

- `ΔM`
- `ΔX/Y/Z`
- `ΔJx/Jy/Jz`

PDF 输出 `mass_properties_plots.pdf` 包含：

- 全箭质量随时间变化曲线。
- 全箭质心 `X/Y/Z` 随时间变化曲线。
- 全箭转动惯量 `Jx/Jy/Jz` 随时间变化曲线。
- 不变质量随时间变化曲线。
- 可变质量汇总随时间变化曲线。
- 可抛质量随时间变化曲线。

## 级间分离规则

不变质量和贮箱可变质量按 QJ 标准随级间/助推器分离发生突变：

```text
t <= ts  -> 保留该级/助推器
t > ts   -> 剔除该级/助推器
```

其中 `ts` 为 `级事件` 表中的分离时间。

## 测试

运行单元测试：

```powershell
python -m pytest tests
```

当前测试覆盖：

- 平行轴定理质量特性合成。
- 圆柱贮箱体积和质心数值积分。
- 等效圆柱贮箱在不填写 QJ 第一类几何时的积分。
- QJ 标准分离边界、`Mef` 工作段消耗、点火前/关机后显式状态。
- 箱底液位结果和运输状态部段合成。

## 当前实现范围和注意事项

- 已实现 QJ 1080A 贮箱半径、推进剂积分、增压气体、可变质量、可抛质量、运输状态和全箭时间序列主流程。
- 已支持 `QJ第一类` 和 `等效圆柱` 两种贮箱几何类型。
- `input_starship_example.xlsx` 是基于公开 Starship/Super Heavy 尺寸和质量数据构造的近似验证样例，不代表 SpaceX 官方内部质量特性模型。
- `贮箱基本参数` 中 `加注后Xf`、`关机后Xr`、`加注后Jyf`、`关机后Jyr` 为空时，程序按 `Mf`/`Mr` 查液位表作为兜底。
- 贮箱几何方程根据 QJ 1080A 公式和图 3/图 4 符号实现。用于真实型号前，建议用 CAD 或已有分析结果校核满箱体积、满箱质心和满箱转动惯量。
- 生成的 Excel/PDF 结果不提交到 Git；需要时可本地重新运行生成。
