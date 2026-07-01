from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

from workbook_formatting import format_output_workbook


EPS = 1e-12


@dataclass(frozen=True)
class MassProperties:
    mass: float = 0.0
    x: float = 0.0
    y: float = 0.0
    z: float = 0.0
    jx: float = 0.0
    jy: float = 0.0
    jz: float = 0.0

    def as_dict(self, prefix: str = "") -> dict[str, float]:
        p = f"{prefix}_" if prefix else ""
        return {
            f"{p}M": self.mass,
            f"{p}X": self.x,
            f"{p}Y": self.y,
            f"{p}Z": self.z,
            f"{p}Jx": self.jx,
            f"{p}Jy": self.jy,
            f"{p}Jz": self.jz,
        }


def _num(value: Any, default: float = 0.0) -> float:
    if value is None or (isinstance(value, float) and np.isnan(value)) or value == "":
        return default
    return float(value)


def _text(value: Any) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    return str(value).strip()


def _has_value(value: Any) -> bool:
    return not (value is None or (isinstance(value, float) and np.isnan(value)) or value == "")


def combine_mass_properties(items: list[MassProperties]) -> MassProperties:
    items = [i for i in items if i.mass > EPS]
    total = sum(i.mass for i in items)
    if total <= EPS:
        return MassProperties()

    x = sum(i.mass * i.x for i in items) / total
    y = sum(i.mass * i.y for i in items) / total
    z = sum(i.mass * i.z for i in items) / total

    jx = sum(i.jx + i.mass * ((y - i.y) ** 2 + (z - i.z) ** 2) for i in items)
    jy = sum(i.jy + i.mass * ((z - i.z) ** 2 + (x - i.x) ** 2) for i in items)
    jz = sum(i.jz + i.mass * ((x - i.x) ** 2 + (y - i.y) ** 2) for i in items)
    return MassProperties(total, x, y, z, jx, jy, jz)


def tank_id(row: pd.Series) -> str:
    return f"{_text(row.get('类型'))}-{int(_num(row.get('所属编号')))}-{int(_num(row.get('贮箱编号')))}"


def tank_radius_type1(x: np.ndarray, p: dict[str, Any]) -> np.ndarray:
    x = np.asarray(x, dtype=float)
    r = np.zeros_like(x)
    slope11 = np.tan(np.deg2rad(_num(p.get("beta11_deg"))))
    slope12 = np.tan(np.deg2rad(_num(p.get("beta12_deg"))))
    semi_a11 = _num(p.get("semi_a11"))
    b11 = _num(p.get("b11"))
    b12 = _num(p.get("b12"))
    mu11 = semi_a11 / b11 if abs(b11) > EPS else 0.0
    mu12 = semi_a11 / b12 if abs(b12) > EPS else 0.0

    segs = [
        ("R11", "0", "H11", lambda xx: _num(p.get("R11"))),
        ("arc12", "H11", "H12", lambda xx: np.sqrt(np.maximum(_num(p.get("R12")) ** 2 - (_num(p.get("h11")) - xx) ** 2, 0))),
        ("cone13", "H12", "H13", lambda xx: _num(p.get("R13")) + slope11 * (xx - _num(p.get("H12")))),
        ("ellipse14", "H13", "H14", lambda xx: mu11 * np.sqrt(np.maximum(b11**2 - (_num(p.get("H14")) - xx) ** 2, 0))),
        ("cyl15", "H14", "H15", lambda xx: semi_a11),
        ("ellipse16", "H15", "H16", lambda xx: mu12 * np.sqrt(np.maximum(b12**2 - (xx - _num(p.get("H15"))) ** 2, 0))),
        ("cone17", "H16", "H17", lambda xx: _num(p.get("R14")) + slope12 * (_num(p.get("H17")) - xx)),
        ("arc18", "H17", "H18", lambda xx: np.sqrt(np.maximum(_num(p.get("R15")) ** 2 - (xx - _num(p.get("h12"))) ** 2, 0))),
        ("R16", "H18", "H19", lambda xx: _num(p.get("R16"))),
    ]
    for _, lo_name, hi_name, fn in segs:
        lo = 0.0 if lo_name == "0" else _num(p.get(lo_name))
        hi = _num(p.get(hi_name))
        if hi <= lo + EPS:
            continue
        mask = (x >= lo - EPS) & (x <= hi + EPS)
        r[mask] = fn(x[mask])
    return np.maximum(r, 0)


def tank_radius_type2(x: np.ndarray, p: dict[str, Any]) -> np.ndarray:
    if _text(p.get("是否启用扣除")) not in {"是", "yes", "YES", "1", "true", "True"}:
        return np.zeros_like(np.asarray(x, dtype=float))
    x = np.asarray(x, dtype=float)
    r = np.zeros_like(x)
    slope21 = np.tan(np.deg2rad(_num(p.get("beta21_deg"))))
    slope22 = np.tan(np.deg2rad(_num(p.get("beta22_deg"))))
    semi_a22 = _num(p.get("semi_a22"))
    b21 = _num(p.get("b21"))
    b22 = _num(p.get("b22"))
    mu21 = semi_a22 / b21 if abs(b21) > EPS else 0.0
    mu22 = semi_a22 / b22 if abs(b22) > EPS else 0.0

    segs = [
        ("R21", "0", "H21", lambda xx: _num(p.get("R21"))),
        ("ellipse22", "H21", "H22", lambda xx: mu21 * np.sqrt(np.maximum(b21**2 - (xx - _num(p.get("H21"))) ** 2, 0))),
        ("cone23", "H22", "H23", lambda xx: _num(p.get("R23")) + slope21 * (_num(p.get("H23")) - xx)),
        ("arc24", "H23", "H24", lambda xx: np.sqrt(np.maximum(_num(p.get("R22")) ** 2 - (xx - _num(p.get("h21"))) ** 2, 0))),
        ("R24", "H24", "H25", lambda xx: _num(p.get("R24"))),
        ("arc26", "H25", "H26", lambda xx: np.sqrt(np.maximum(_num(p.get("R25")) ** 2 - (_num(p.get("h22")) - xx) ** 2, 0))),
        ("cone27", "H26", "H27", lambda xx: _num(p.get("R26")) + slope22 * (xx - _num(p.get("H26")))),
        ("ellipse28", "H27", "H28", lambda xx: mu22 * np.sqrt(np.maximum(b22**2 - (_num(p.get("H28")) - xx) ** 2, 0))),
        ("R27", "H28", "H29", lambda xx: _num(p.get("R27"))),
    ]
    for _, lo_name, hi_name, fn in segs:
        lo = 0.0 if lo_name == "0" else _num(p.get(lo_name))
        hi = _num(p.get(hi_name))
        if hi <= lo + EPS:
            continue
        mask = (x >= lo - EPS) & (x <= hi + EPS)
        r[mask] = fn(x[mask])
    return np.maximum(r, 0)


def tank_area(x: np.ndarray, p1: dict[str, Any], p2: dict[str, Any], tank: dict[str, Any] | None = None) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    if tank and _text(tank.get("几何类型")) == "等效圆柱":
        r = max(_num(tank.get("等效半径R")), 0.0)
        r1 = np.full_like(np.asarray(x, dtype=float), r)
        r2 = np.zeros_like(r1)
        return np.pi * r1**2, r1, r2
    r1 = tank_radius_type1(x, p1)
    r2 = tank_radius_type2(x, p2)
    area = np.pi * np.maximum(r1**2 - r2**2, 0)
    return area, r1, r2


def _trapz(y: np.ndarray, x: np.ndarray) -> float:
    return float(np.trapz(y, x)) if len(x) > 1 else 0.0


def propellant_by_level(tank: dict[str, Any], p1: dict[str, Any], p2: dict[str, Any], ducts: pd.DataFrame, n: int) -> pd.DataFrame:
    if _text(tank.get("几何类型")) == "等效圆柱":
        height = _num(tank.get("等效高度H"))
    else:
        height = max(_num(p1.get("H19")), _num(p2.get("H29")))
    if height <= EPS:
        height = _num(tank.get("总高H"), 1.0)
    rho = _num(tank.get("推进剂密度rho_phi"))
    rows = []
    h_values = np.linspace(0, height, max(2, int(n)))
    tid = tank["贮箱ID"]
    duct_items = []
    for _, d in ducts.iterrows():
        duct_items.append(MassProperties(_num(d.get("m_phi")), 0, _num(d.get("Y_phi")), _num(d.get("Z_phi")), _num(d.get("Jx_phi")), 0, 0))
    duct_mp = combine_mass_properties(duct_items)

    for h in h_values:
        x = np.linspace(0, h, max(2, int(n))) if h > EPS else np.array([0.0, 0.0])
        area, r1, r2 = tank_area(x, p1, p2, tank)
        v = _trapz(area, x)
        m = rho * v
        xcg = _trapz(x * area, x) / v if v > EPS else 0.0
        integrand = np.pi * x**2 * np.maximum(r1**2 - r2**2, 0) + np.pi / 4 * np.maximum(r1**4 - r2**4, 0)
        jy = rho * (_trapz(integrand, x) - xcg**2 * v) if v > EPS else 0.0
        rows.append({
            "贮箱ID": tid, "h": h, "V_phi": v, "M_phi": m, "X_phi": xcg,
            "Y_phi": duct_mp.y if duct_mp.mass > EPS else 0.0,
            "Z_phi": duct_mp.z if duct_mp.mass > EPS else 0.0,
            "Jx_phi": duct_mp.jx, "Jy_phi": max(jy, 0.0), "Jz_phi": max(jy, 0.0),
        })
    return pd.DataFrame(rows)


def gas_and_variable_by_level(tank: dict[str, Any], prop: pd.DataFrame) -> pd.DataFrame:
    rho_g = _num(tank.get("气体密度rho_g"))
    rho_phi = max(_num(tank.get("推进剂密度rho_phi")), EPS)
    vo = _num(tank.get("总容积Vo"), float(prop["V_phi"].max()))
    xo = _num(tank.get("满箱Xo"), float(prop.loc[prop["V_phi"].idxmax(), "X_phi"]))
    jyo = _num(tank.get("满箱Jyo"), float(prop["Jy_phi"].max()))
    out = prop.copy()
    out["V_g"] = np.maximum(vo - out["V_phi"], 0.0)
    out["M_g"] = rho_g * out["V_g"]
    out["X_g"] = np.where(out["V_g"] > EPS, (vo * xo - out["V_phi"] * out["X_phi"]) / out["V_g"], 0.0)
    out["Y_g"] = 0.0
    out["Z_g"] = 0.0
    out["Jx_g"] = 0.0
    out["Jy_g"] = (jyo - out["Jy_phi"]) * rho_g / rho_phi - rho_g * out["V_g"] * (xo - out["X_g"]) ** 2 - rho_g * out["V_phi"] * (xo - out["X_phi"]) ** 2
    out["Jy_g"] = out["Jy_g"].clip(lower=0.0)
    out["Jz_g"] = out["Jy_g"]
    out["M_v"] = out["M_phi"] + out["M_g"]
    out["X_v"] = np.where(out["M_v"] > EPS, (out["M_phi"] * out["X_phi"] + out["M_g"] * out["X_g"]) / out["M_v"], 0.0)
    out["Y_v"] = out["Y_phi"]
    out["Z_v"] = out["Z_phi"]
    out["Jx_v"] = out["Jx_phi"]
    out["Jy_v"] = out["Jy_phi"] + out["Jy_g"] + out["M_phi"] * (out["X_v"] - out["X_phi"]) ** 2 + out["M_g"] * (out["X_v"] - out["X_g"]) ** 2
    out["Jz_v"] = out["Jy_v"]
    return out


def mp_from_row(row: pd.Series | dict[str, Any], prefix: str = "") -> MassProperties:
    def key(name: str) -> str:
        return f"{prefix}{name}" if prefix else name
    return MassProperties(_num(row.get(key("M"))), _num(row.get(key("X"))), _num(row.get(key("Y"))), _num(row.get(key("Z"))), _num(row.get(key("Jx"))), _num(row.get(key("Jy"))), _num(row.get(key("Jz"))))


def read_input_excel(path: str | Path) -> dict[str, pd.DataFrame]:
    sheets = pd.read_excel(path, sheet_name=None)
    return {k: v.dropna(how="all") for k, v in sheets.items()}


def validate_inputs(data: dict[str, pd.DataFrame]) -> pd.DataFrame:
    warnings = []
    required = ["总体参数", "级事件", "不变质量", "可抛质量", "贮箱基本参数", "贮箱几何_第一类", "贮箱几何_第二类", "外行导管"]
    for s in required:
        if s not in data:
            warnings.append({"级别": "错误", "位置": s, "说明": "缺少工作表"})
    if "级事件" in data:
        for i, r in data["级事件"].iterrows():
            if _num(r.get("关机时间tb")) < _num(r.get("点火时间tf")) or _num(r.get("分离时间ts")) < _num(r.get("关机时间tb")):
                warnings.append({"级别": "错误", "位置": f"级事件 第{i+2}行", "说明": "时间顺序应满足 tf <= tb <= ts"})
    for sheet in ["不变质量", "可抛质量", "贮箱基本参数"]:
        if sheet in data:
            for i, r in data[sheet].iterrows():
                for c in r.index:
                    if any(k in c for k in ["质量", "密度", "秒流量", "M", "rho", "mdot"]):
                        if _num(r.get(c)) < 0:
                            warnings.append({"级别": "错误", "位置": f"{sheet} 第{i+2}行 {c}", "说明": "数值不能为负"})
    return pd.DataFrame(warnings or [{"级别": "通过", "位置": "全部", "说明": "未发现阻断性输入错误"}])


def invariant_results(data: dict[str, pd.DataFrame]) -> pd.DataFrame:
    df = data.get("不变质量", pd.DataFrame())
    rows = []
    for (typ, owner), g in df.groupby(["类型", "所属编号"], dropna=True):
        mp = combine_mass_properties([mp_from_row(r) for _, r in g.iterrows()])
        rows.append({"类型": typ, "编号": int(owner), **mp.as_dict()})
    return pd.DataFrame(rows)


def jettison_at_time(t: float, df: pd.DataFrame) -> MassProperties:
    if df.empty:
        return MassProperties()
    items = []
    for _, r in df.iterrows():
        if t <= _num(r.get("抛离时间")):
            items.append(mp_from_row(r))
    return combine_mass_properties(items)


def invariant_mp_at_time(t: float, inv: pd.DataFrame, events: pd.DataFrame) -> MassProperties:
    if inv.empty:
        return MassProperties()
    items = []
    for _, r in inv.iterrows():
        typ = _text(r.get("类型"))
        idx = int(_num(r.get("编号")))
        event = events[(events["类型"].astype(str).str.strip() == typ) & (events["编号"].astype(float) == idx)] if not events.empty else pd.DataFrame()
        if event.empty:
            keep = True
        else:
            ts = _num(event.iloc[0].get("分离时间ts"), float("inf"))
            keep = t <= ts
        if keep:
            items.append(mp_from_row({"M": r["M"], "X": r["X"], "Y": r["Y"], "Z": r["Z"], "Jx": r["Jx"], "Jy": r["Jy"], "Jz": r["Jz"]}))
    return combine_mass_properties(items)


def build_tank_level_results(data: dict[str, pd.DataFrame], n: int) -> pd.DataFrame:
    tanks = data.get("贮箱基本参数", pd.DataFrame()).copy()
    if tanks.empty:
        return pd.DataFrame()
    geom1 = data.get("贮箱几何_第一类", pd.DataFrame()).set_index("贮箱ID", drop=False)
    geom2 = data.get("贮箱几何_第二类", pd.DataFrame()).set_index("贮箱ID", drop=False)
    ducts = data.get("外行导管", pd.DataFrame())
    all_rows = []
    for _, row in tanks.iterrows():
        row = row.copy()
        tid = tank_id(row)
        row["贮箱ID"] = tid
        p1 = geom1.loc[tid].to_dict() if tid in geom1.index else {}
        p2 = geom2.loc[tid].to_dict() if tid in geom2.index else {"是否启用扣除": "否"}
        duct_rows = ducts[ducts["贮箱ID"] == tid] if not ducts.empty and "贮箱ID" in ducts else pd.DataFrame()
        prop = propellant_by_level(row.to_dict(), p1, p2, duct_rows, n)
        all_rows.append(gas_and_variable_by_level(row.to_dict(), prop))
    return pd.concat(all_rows, ignore_index=True) if all_rows else pd.DataFrame()


def build_tank_bottom_level_results(data: dict[str, pd.DataFrame], level: pd.DataFrame) -> pd.DataFrame:
    tanks = data.get("贮箱基本参数", pd.DataFrame()).copy()
    if tanks.empty or level.empty:
        return pd.DataFrame()
    rows = []
    for _, tank in tanks.iterrows():
        tid = tank_id(tank)
        sub = level[level["贮箱ID"] == tid].sort_values("h").reset_index(drop=True)
        if sub.empty:
            continue
        hb = _num(tank.get("箱底hb"))
        mdot = _num(tank.get("秒流量mdot"))
        mef = _num(tank.get("启动后剩余量Mef"), _num(tank.get("加注量Mf")))
        base = _nearest_by_column(sub, "h", hb)
        base_v = _num(base.get("V_phi"))
        base_m = _num(base.get("M_phi"))
        base_x = _num(base.get("X_phi"))
        base_jy = _num(base.get("Jy_phi"))
        for _, r in sub.iterrows():
            h = _num(r.get("h"))
            if h < hb - EPS:
                continue
            v_b = max(_num(r.get("V_phi")) - base_v, 0.0)
            m_b = max(_num(r.get("M_phi")) - base_m, 0.0)
            x_phi = _num(r.get("X_phi"))
            x_b = (_num(r.get("M_phi")) * x_phi - base_m * base_x) / m_b - hb if m_b > EPS else 0.0
            jy_b = _num(r.get("Jy_phi")) - base_jy - base_m * (x_phi - base_x) ** 2 - m_b * (x_phi - x_b - hb) ** 2
            t_b = (mef - m_b - base_m) / mdot if mdot > EPS else 0.0
            rows.append({
                "贮箱ID": tid,
                "h": h,
                "hb": hb,
                "h_b": h - hb,
                "V_b": v_b,
                "M_b": m_b,
                "X_b": x_b,
                "Jy_b": max(jy_b, 0.0),
                "Jz_b": max(jy_b, 0.0),
                "T_b": t_b,
            })
    return pd.DataFrame(rows)


def _nearest_by_column(df: pd.DataFrame, column: str, value: float) -> pd.Series:
    return df.loc[(df[column] - value).abs().idxmin()]


def _mp_from_level_row(row: pd.Series, mass: float | None = None, x: float | None = None, jy: float | None = None) -> MassProperties:
    m = _num(row.get("M_v")) if mass is None else mass
    xx = _num(row.get("X_v")) if x is None else x
    yy = _num(row.get("Y_v"))
    zz = _num(row.get("Z_v"))
    jx = _num(row.get("Jx_v"))
    jyy = _num(row.get("Jy_v")) if jy is None else jy
    return MassProperties(m, xx, yy, zz, jx, jyy, jyy)


def tank_mp_at_time(t: float, tank: pd.Series, level: pd.DataFrame) -> MassProperties:
    tf, tb, ts = _num(tank.get("点火时间tf")), _num(tank.get("关机时间tb")), _num(tank.get("分离时间ts"))
    if t > ts:
        return MassProperties()
    mdot = _num(tank.get("秒流量mdot"))
    mf = _num(tank.get("加注量Mf"))
    mr = _num(tank.get("关机剩余量Mr"))
    if level.empty:
        return MassProperties()
    mef = _num(tank.get("启动后剩余量Mef"), mf)
    if t < tf:
        target_m = mf
        r = _nearest_by_column(level, "M_v", target_m)
        local_mp = _mp_from_level_row(
            r,
            mass=target_m,
            x=_num(tank.get("加注后Xf")) if _has_value(tank.get("加注后Xf")) else None,
            jy=_num(tank.get("加注后Jyf")) if _has_value(tank.get("加注后Jyf")) else None,
        )
    elif t <= tb:
        target_m = max(mr, mef - mdot * (t - tf))
        r = _nearest_by_column(level, "M_v", target_m)
        local_mp = _mp_from_level_row(r, mass=target_m)
    else:
        target_m = mr
        r = _nearest_by_column(level, "M_v", target_m)
        local_mp = _mp_from_level_row(
            r,
            mass=target_m,
            x=_num(tank.get("关机后Xr")) if _has_value(tank.get("关机后Xr")) else None,
            jy=_num(tank.get("关机后Jyr")) if _has_value(tank.get("关机后Jyr")) else None,
        )
    is_booster = _text(tank.get("类型")) == "助推器"
    x0 = abs(_num(tank.get("坐标原点X或Loki")))
    x = x0 - local_mp.x
    y = _num(tank.get("坐标原点Y")) + local_mp.y if is_booster else local_mp.y
    z = _num(tank.get("坐标原点Z")) + local_mp.z if is_booster else local_mp.z
    return MassProperties(local_mp.mass, x, y, z, local_mp.jx, local_mp.jy, local_mp.jz)


def time_series_results(data: dict[str, pd.DataFrame], inv: pd.DataFrame, level: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    params = data["总体参数"].set_index("参数")["值"]
    t0, t1, dt = _num(params.get("飞行开始时间")), _num(params.get("飞行结束时间")), _num(params.get("时间步长"), 1.0)
    times = np.arange(t0, t1 + 0.5 * dt, dt)
    tanks = data.get("贮箱基本参数", pd.DataFrame()).copy()
    if not tanks.empty:
        tanks["贮箱ID"] = tanks.apply(tank_id, axis=1)
    jet = data.get("可抛质量", pd.DataFrame())
    events = data.get("级事件", pd.DataFrame())

    inv_rows, tank_rows, jet_rows, total_rows = [], [], [], []
    for t in times:
        inv_mp = invariant_mp_at_time(float(t), inv, events)
        inv_rows.append({"t": t, **inv_mp.as_dict()})
        variable = []
        for _, tank in tanks.iterrows():
            sub = level[level["贮箱ID"] == tank["贮箱ID"]] if not level.empty else pd.DataFrame()
            mp = tank_mp_at_time(float(t), tank, sub)
            variable.append(mp)
            tank_rows.append({"t": t, "贮箱ID": tank["贮箱ID"], **mp.as_dict()})
        jmp = jettison_at_time(float(t), jet)
        jet_rows.append({"t": t, **jmp.as_dict()})
        total = combine_mass_properties([inv_mp, jmp, *variable])
        total_rows.append({"t": t, **total.as_dict()})
    return pd.DataFrame(inv_rows), pd.DataFrame(tank_rows), pd.DataFrame(jet_rows), pd.DataFrame(total_rows)


def _nearest_row(df: pd.DataFrame, t: float) -> pd.Series:
    return df.loc[(df["t"] - t).abs().idxmin()]


def build_key_time_summary(data: dict[str, pd.DataFrame], total: pd.DataFrame) -> pd.DataFrame:
    events = [{"事件类型": "飞行开始", "对象": "全箭", "事件时间": float(total["t"].min()), "说明": "时间序列起点"}, {"事件类型": "飞行结束", "对象": "全箭", "事件时间": float(total["t"].max()), "说明": "时间序列终点"}]
    for _, r in data.get("级事件", pd.DataFrame()).iterrows():
        obj = f"{_text(r.get('类型'))}{int(_num(r.get('编号')))}"
        events.extend([
            {"事件类型": "点火", "对象": obj, "事件时间": _num(r.get("点火时间tf")), "说明": "级事件表"},
            {"事件类型": "关机", "对象": obj, "事件时间": _num(r.get("关机时间tb")), "说明": "级事件表"},
            {"事件类型": "分离", "对象": obj, "事件时间": _num(r.get("分离时间ts")), "说明": "t > ts 后剔除"},
        ])
    for _, r in data.get("可抛质量", pd.DataFrame()).iterrows():
        if "抛离时间" in r and not pd.isna(r.get("抛离时间")):
            events.append({"事件类型": "可抛质量抛离", "对象": _text(r.get("名称")), "事件时间": _num(r.get("抛离时间")), "说明": "可抛质量表"})

    rows = []
    for e in events:
        tr = _nearest_row(total, e["事件时间"])
        rows.append({**e, "匹配时间": tr["t"], "M": tr["M"], "X": tr["X"], "Y": tr["Y"], "Z": tr["Z"], "Jx": tr["Jx"], "Jy": tr["Jy"], "Jz": tr["Jz"]})
    return pd.DataFrame(rows).drop_duplicates(subset=["事件类型", "对象", "事件时间"]).sort_values(["事件时间", "事件类型", "对象"]).reset_index(drop=True)


def build_event_delta_summary(data: dict[str, pd.DataFrame], total: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, r in data.get("级事件", pd.DataFrame()).iterrows():
        obj = f"{_text(r.get('类型'))}{int(_num(r.get('编号')))}"
        for event_name, col in [("点火", "点火时间tf"), ("关机", "关机时间tb"), ("分离", "分离时间ts")]:
            t = _num(r.get(col))
            before = total[total["t"] < t]
            after = total[total["t"] >= t]
            if before.empty or after.empty:
                continue
            b = before.iloc[-1]
            a = after.iloc[0]
            rows.append({
                "事件类型": event_name,
                "对象": obj,
                "事件时间": t,
                "前一时间": b["t"],
                "后一时间": a["t"],
                "ΔM": a["M"] - b["M"],
                "ΔX": a["X"] - b["X"],
                "ΔY": a["Y"] - b["Y"],
                "ΔZ": a["Z"] - b["Z"],
                "ΔJx": a["Jx"] - b["Jx"],
                "ΔJy": a["Jy"] - b["Jy"],
                "ΔJz": a["Jz"] - b["Jz"],
            })
    return pd.DataFrame(rows)


def transport_results(data: dict[str, pd.DataFrame]) -> pd.DataFrame:
    df = data.get("运输状态", pd.DataFrame())
    if df.empty:
        return pd.DataFrame()
    rows = []
    for section_id, g in df.groupby("部段编号", dropna=True):
        items = []
        for _, r in g.iterrows():
            items.append(MassProperties(
                _num(r.get("M")),
                _num(r.get("Xcq")),
                _num(r.get("Ycq")),
                _num(r.get("Zcq")),
                _num(r.get("Jxcq")),
                _num(r.get("Jycq")),
                _num(r.get("Jzcq")),
            ))
        mp = combine_mass_properties(items)
        first = g.iloc[0]
        rows.append({
            "部段编号": int(_num(section_id)),
            "部段名称": _text(first.get("部段名称")),
            "Xq": _num(first.get("Xq")),
            "M_cq": mp.mass,
            "X_cq": mp.x,
            "Y_cq": mp.y,
            "Z_cq": mp.z,
            "Jx_cq": mp.jx,
            "Jy_cq": mp.jy,
            "Jz_cq": mp.jz,
        })
    return pd.DataFrame(rows)


def validate_computed_inputs(data: dict[str, pd.DataFrame], level: pd.DataFrame) -> pd.DataFrame:
    warnings = []
    tanks = data.get("贮箱基本参数", pd.DataFrame())
    if tanks.empty:
        return pd.DataFrame()
    for _, tank in tanks.iterrows():
        tid = tank_id(tank)
        mf = _num(tank.get("加注量Mf"))
        mef = _num(tank.get("启动后剩余量Mef"), mf)
        mr = _num(tank.get("关机剩余量Mr"))
        if mef > mf + EPS:
            warnings.append({"级别": "警告", "位置": f"贮箱基本参数 {tid}", "说明": "启动后剩余量Mef 不应大于 加注量Mf"})
        if mr > mef + EPS:
            warnings.append({"级别": "警告", "位置": f"贮箱基本参数 {tid}", "说明": "关机剩余量Mr 不应大于 启动后剩余量Mef"})
        sub = level[level["贮箱ID"] == tid] if not level.empty and "贮箱ID" in level else pd.DataFrame()
        if sub.empty:
            continue
        min_m, max_m = float(sub["M_v"].min()), float(sub["M_v"].max())
        for name, value in [("加注量Mf", mf), ("启动后剩余量Mef", mef), ("关机剩余量Mr", mr)]:
            if value < min_m - EPS or value > max_m + EPS:
                warnings.append({"级别": "警告", "位置": f"贮箱基本参数 {tid} {name}", "说明": f"质量 {value} 超出液位表 M_v 范围 {min_m} ~ {max_m}，将使用最近液位的质心/惯量"})
        if _has_value(tank.get("总容积Vo")):
            vo = _num(tank.get("总容积Vo"))
            vmax = float(sub["V_phi"].max())
            if vmax > EPS and abs(vo - vmax) / vmax > 0.01:
                warnings.append({"级别": "警告", "位置": f"贮箱基本参数 {tid} 总容积Vo", "说明": f"Vo 与几何积分最大体积偏差超过 1%：Vo={vo}, V_phi_max={vmax}"})
    return pd.DataFrame(warnings)


def build_result_overview(data: dict[str, pd.DataFrame], results: dict[str, pd.DataFrame], input_path: str | Path) -> pd.DataFrame:
    total = results["全箭时间序列"]
    initial = total.iloc[0]
    final = total.iloc[-1]
    nonzero = total[total["M"] > EPS]
    rows = [
        ["输入文件", str(input_path), ""],
        ["时间范围", f"{total['t'].min()} ~ {total['t'].max()}", "s"],
        ["时间点数量", len(total), ""],
        ["贮箱数量", len(data.get("贮箱基本参数", pd.DataFrame())), ""],
        ["级事件数量", len(data.get("级事件", pd.DataFrame())), ""],
        ["初始全箭质量", initial["M"], "kg"],
        ["初始质心X", initial["X"], "m"],
        ["初始质心Y", initial["Y"], "m"],
        ["初始质心Z", initial["Z"], "m"],
        ["初始Jx", initial["Jx"], "kg·m^2"],
        ["初始Jy", initial["Jy"], "kg·m^2"],
        ["初始Jz", initial["Jz"], "kg·m^2"],
        ["最大质量", total["M"].max(), "kg"],
        ["最小非零质量", nonzero["M"].min() if not nonzero.empty else 0, "kg"],
        ["最终质量", final["M"], "kg"],
        ["输入检查", _text(results["输入检查"].iloc[0].get("级别")) if not results["输入检查"].empty else "", ""],
    ]
    return pd.DataFrame(rows, columns=["项目", "值", "单位"])


def calculate(input_path: str | Path) -> dict[str, pd.DataFrame]:
    data = read_input_excel(input_path)
    checks = validate_inputs(data)
    params = data["总体参数"].set_index("参数")["值"]
    n = int(_num(params.get("液位积分点数"), 1001))
    inv = invariant_results(data)
    level = build_tank_level_results(data, n)
    computed_checks = validate_computed_inputs(data, level)
    if not computed_checks.empty:
        checks = pd.concat([checks, computed_checks], ignore_index=True)
    bottom_level = build_tank_bottom_level_results(data, level)
    transport = transport_results(data)
    invariant_ts, variable_ts, jet_ts, total_ts = time_series_results(data, inv, level)
    results = {
        "输入检查": checks,
        "不变质量结果": inv,
        "贮箱液位结果": level,
        "贮箱箱底液位结果": bottom_level,
        "运输状态结果": transport,
        "不变质量时间序列": invariant_ts,
        "可变质量时间序列": variable_ts,
        "可抛质量时间序列": jet_ts,
        "全箭时间序列": total_ts,
    }
    key_summary = build_key_time_summary(data, total_ts)
    delta_summary = build_event_delta_summary(data, total_ts)
    overview = build_result_overview(data, results, input_path)
    return {"结果总览": overview, "关键时间点汇总": key_summary, "关键事件前后对比": delta_summary, **results}


def write_output(results: dict[str, pd.DataFrame], output_path: str | Path) -> None:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for name, df in results.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    format_output_workbook(str(output_path))


def write_plots(results: dict[str, pd.DataFrame], output_path: str | Path) -> None:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages
    from matplotlib import font_manager, rcParams

    for font in ["Microsoft YaHei", "SimHei", "Arial Unicode MS", "DejaVu Sans"]:
        if any(f.name == font for f in font_manager.fontManager.ttflist):
            rcParams["font.sans-serif"] = [font]
            break
    rcParams["axes.unicode_minus"] = False

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    total = results.get("全箭时间序列", pd.DataFrame())
    invariant = results.get("不变质量时间序列", pd.DataFrame())
    variable = results.get("可变质量时间序列", pd.DataFrame())
    jettison = results.get("可抛质量时间序列", pd.DataFrame())

    def page(title: str, df: pd.DataFrame, cols: list[str], ylabel: str) -> None:
        if df.empty or "t" not in df:
            return
        fig, ax = plt.subplots(figsize=(10.5, 6.2))
        for col in cols:
            if col in df:
                ax.plot(df["t"], df[col], label=col, linewidth=1.8)
        ax.set_title(title)
        ax.set_xlabel("t / s")
        ax.set_ylabel(ylabel)
        ax.grid(True, alpha=0.3)
        ax.legend()
        fig.tight_layout()
        pdf.savefig(fig)
        plt.close(fig)

    with PdfPages(output_path) as pdf:
        page("全箭质量随时间变化", total, ["M"], "M / kg")
        page("全箭质心随时间变化", total, ["X", "Y", "Z"], "Coordinate / m")
        page("全箭转动惯量随时间变化", total, ["Jx", "Jy", "Jz"], "J / kg m^2")
        page("不变质量随时间变化", invariant, ["M"], "M_u / kg")

        if not variable.empty:
            var_sum = variable.groupby("t", as_index=False)["M"].sum()
            page("可变质量汇总随时间变化", var_sum, ["M"], "sum(M_v) / kg")
        page("可抛质量随时间变化", jettison, ["M"], "M_thw / kg")
