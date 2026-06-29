# Rocket CG Calculator

Liquid rocket mass, center of gravity, and moment of inertia calculator based on **QJ 1080A-1997 液体火箭质量、质心和转动惯量计算方法**.

The project uses Excel only for engineering inputs and result delivery. The calculation logic is implemented in Python for better transparency, testing, and maintainability.

## Features

- Read structured rocket mass-property inputs from `input_template.xlsx`.
- Calculate component mass-property combinations using the parallel-axis theorem.
- Model invariant mass changes at stage/booster separation times.
- Calculate jettisoned mass time histories.
- Implement QJ 1080A tank geometry radius models for propellant/ullage calculations.
- Numerically integrate propellant volume, center of gravity, and inertia versus liquid level.
- Calculate variable tank mass properties over flight time.
- Output Excel result tables and a PDF plot report.
- Include pytest coverage for core calculations.

## Repository Contents

```text
rocket_mass_properties.py   Core calculation library
run_calculation.py          Command-line entry point
create_input_template.py    Regenerate the Excel input template
input_template.xlsx         Editable input workbook template
```

Generated outputs are written under `output/` and are intentionally ignored by Git.

## Requirements

Python 3.10+ is recommended.

Required packages:

```text
numpy
pandas
openpyxl
matplotlib
pytest
```

Install dependencies with:

```powershell
pip install numpy pandas openpyxl matplotlib pytest
```

## Usage

Create or refresh the input template:

```powershell
python create_input_template.py
```

Edit `input_template.xlsx`, then run the calculation:

```powershell
python run_calculation.py
```

Default outputs:

```text
output/mass_properties_result.xlsx
output/mass_properties_plots.pdf
```

Specify custom paths:

```powershell
python run_calculation.py --input input_template.xlsx --output output/mass_properties_result.xlsx --plots output/mass_properties_plots.pdf
```

Skip PDF generation:

```powershell
python run_calculation.py --no-plots
```

## Input Workbook Sheets

`input_template.xlsx` contains these sheets:

- `总体参数`: stage count, booster count, time range, time step, liquid-level integration points.
- `级事件`: ignition, cutoff, and separation times for stages and boosters.
- `不变质量`: dry/invariant component mass, CG, and inertia data.
- `可抛质量`: fairings or other jettisoned components.
- `贮箱基本参数`: tank density, flow rate, fill mass, residual mass, coordinate offsets, and reference full-tank values.
- `贮箱几何_第一类`: QJ 1080A type-1 tank radius parameters.
- `贮箱几何_第二类`: QJ 1080A type-2 internal deduction radius parameters.
- `外行导管`: off-axis pipe contributions to `Jx`.
- `运输状态`: reserved transport-state inputs.

Angles such as `beta11_deg` are entered in degrees. Python converts them to radians internally.

## Output Workbook Sheets

`mass_properties_result.xlsx` includes:

- `输入检查`
- `不变质量结果`
- `贮箱液位结果`
- `不变质量时间序列`
- `可变质量时间序列`
- `可抛质量时间序列`
- `全箭时间序列`

## Separation Rule

Invariant mass follows this stage/booster separation rule:

```text
t < ts   -> retained
t >= ts  -> removed
```

This matches the implemented boundary convention for stage separation events.

## Tests

Run tests with:

```powershell
python -m pytest tests
```

Current tests cover:

- Parallel-axis mass-property combination.
- Cylindrical tank numerical volume/CG integration.
- Invariant mass removal at separation time.

## Current Scope and Notes

- Core QJ 1080A tank radius, propellant, gas, variable mass, jettison, and total rocket time-series workflows are implemented.
- `运输状态` input is currently reserved and not yet connected to the output workflow.
- The tank geometry equations are implemented from QJ 1080A formulas and the provided figure references. Real vehicle use should validate full-tank volume, CG, and inertia against known CAD or analysis data.
- Generated Excel/PDF outputs are not committed; regenerate them locally when needed.
