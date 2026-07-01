from math import pi
from pathlib import Path

import pandas as pd

from workbook_formatting import format_input_workbook


OUT = Path("input_starship_example.xlsx")
RADIUS = 4.5
LOX_RHO = 1141.0
CH4_RHO = 422.0


def cyl_height(mass: float, rho: float, radius: float = RADIUS) -> float:
    return mass / rho / (pi * radius**2)


def cyl_jx(mass: float, radius: float = RADIUS) -> float:
    return 0.5 * mass * radius**2


def cyl_jy(mass: float, length: float, radius: float = RADIUS) -> float:
    return mass * (3 * radius**2 + length**2) / 12


def tank_row(stage: int, tank_no: int, name: str, prop_mass: float, rho: float, tf: float, tb: float, ts: float, bottom_x: float, burn_time: float) -> list:
    h = cyl_height(prop_mass, rho)
    vo = pi * RADIUS**2 * h
    mdot = prop_mass / burn_time
    return [
        "子级", stage, tank_no, name, "等效圆柱", RADIUS, h,
        rho, 5.0, mdot, prop_mass, 0.0, prop_mass,
        h / 2, 0.0, cyl_jy(prop_mass, h), 0.0,
        bottom_x, 0.0, 0.0, tf, tb, ts, 0.0, h, vo, h / 2, cyl_jy(prop_mass, h),
        "公开数据近似验证样例，不代表真实星舰内部设计",
    ]


def write_example(path: Path = OUT) -> None:
    super_heavy_h = 71.0
    ship_h = 52.1
    total_h = super_heavy_h + ship_h
    sh_dry = 275_000.0
    ship_dry = 85_000.0
    sh_lox = 2_700_000.0
    sh_ch4 = 700_000.0
    ship_lox = 1_170_000.0
    ship_ch4 = 330_000.0
    sh_tf, sh_tb, sh_ts = 0.0, 170.0, 170.0
    ship_tf, ship_tb, ship_ts = 170.0, 520.0, 900.0

    sheets = {
        "总体参数": pd.DataFrame([
            ["火箭总级数Nr", 2, "-", "Starship + Super Heavy 近似两级样例"],
            ["助推器数量Nzt", 0, "-", ""],
            ["飞行开始时间", 0, "s", ""],
            ["飞行结束时间", 900, "s", ""],
            ["时间步长", 5, "s", ""],
            ["液位积分点数", 501, "-", "验证样例；正式分析可提高"],
        ], columns=["参数", "值", "单位", "说明"]),
        "级事件": pd.DataFrame([
            ["子级", 1, sh_tf, sh_tb, sh_ts, "Super Heavy 近似"],
            ["子级", 2, ship_tf, ship_tb, ship_ts, "Starship 上面级近似"],
        ], columns=["类型", "编号", "点火时间tf", "关机时间tb", "分离时间ts", "备注"]),
        "不变质量": pd.DataFrame([
            ["子级", 1, "Super Heavy 干质量", sh_dry, ship_h + super_heavy_h / 2, 0, 0, cyl_jx(sh_dry), cyl_jy(sh_dry, super_heavy_h), cyl_jy(sh_dry, super_heavy_h), "公开数据近似"],
            ["子级", 2, "Starship 干质量", ship_dry, ship_h / 2, 0, 0, cyl_jx(ship_dry), cyl_jy(ship_dry, ship_h), cyl_jy(ship_dry, ship_h), "公开数据近似"],
        ], columns=["类型", "所属编号", "组件名称", "M", "X", "Y", "Z", "Jx", "Jy", "Jz", "备注"]),
        "可抛质量": pd.DataFrame(columns=["名称", "M", "X", "Y", "Z", "Jx", "Jy", "Jz", "抛离时间", "备注"]),
        "贮箱基本参数": pd.DataFrame([
            tank_row(1, 1, "Super Heavy LOX", sh_lox, LOX_RHO, sh_tf, sh_tb, sh_ts, total_h, sh_tb - sh_tf),
            tank_row(1, 2, "Super Heavy CH4", sh_ch4, CH4_RHO, sh_tf, sh_tb, sh_ts, total_h - cyl_height(sh_lox, LOX_RHO), sh_tb - sh_tf),
            tank_row(2, 1, "Starship LOX", ship_lox, LOX_RHO, ship_tf, ship_tb, ship_ts, ship_h, ship_tb - ship_tf),
            tank_row(2, 2, "Starship CH4", ship_ch4, CH4_RHO, ship_tf, ship_tb, ship_ts, ship_h - cyl_height(ship_lox, LOX_RHO), ship_tb - ship_tf),
        ], columns=["类型", "所属编号", "贮箱编号", "名称", "几何类型", "等效半径R", "等效高度H", "推进剂密度rho_phi", "气体密度rho_g", "秒流量mdot", "加注量Mf", "关机剩余量Mr", "启动后剩余量Mef", "加注后Xf", "关机后Xr", "加注后Jyf", "关机后Jyr", "坐标原点X或Loki", "坐标原点Y", "坐标原点Z", "点火时间tf", "关机时间tb", "分离时间ts", "箱底hb", "总高H", "总容积Vo", "满箱Xo", "满箱Jyo", "备注"]),
        "贮箱几何_第一类": pd.DataFrame(columns=["贮箱ID", "R11", "R12", "R13", "R14", "R15", "R16", "H11", "H12", "H13", "H14", "H15", "H16", "H17", "H18", "H19", "h11", "h12", "semi_a11", "b11", "b12", "beta11_deg", "beta12_deg", "备注"]),
        "贮箱几何_第二类": pd.DataFrame(columns=["贮箱ID", "是否启用扣除", "R21", "R22", "R23", "R24", "R25", "R26", "R27", "H21", "H22", "H23", "H24", "H25", "H26", "H27", "H28", "H29", "h21", "h22", "semi_a22", "b21", "b22", "beta21_deg", "beta22_deg", "备注"]),
        "外行导管": pd.DataFrame(columns=["贮箱ID", "导管编号", "m_phi", "Y_phi", "Z_phi", "Jx_phi", "备注"]),
        "运输状态": pd.DataFrame(columns=["部段编号", "部段名称", "Xq", "组件名称", "M", "Xcq", "Ycq", "Zcq", "Jxcq", "Jycq", "Jzcq", "备注"]),
    }

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)

    format_input_workbook(str(path))


if __name__ == "__main__":
    write_example()
    print(f"saved {OUT.resolve()}")
