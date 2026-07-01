import math

import numpy as np
import pandas as pd

from rocket_mass_properties import MassProperties, build_tank_bottom_level_results, combine_mass_properties, invariant_mp_at_time, propellant_by_level, tank_mp_at_time, transport_results


def test_combine_two_point_masses():
    result = combine_mass_properties([
        MassProperties(1, 0, 0, 0, 0, 0, 0),
        MassProperties(1, 2, 0, 0, 0, 0, 0),
    ])
    assert math.isclose(result.mass, 2)
    assert math.isclose(result.x, 1)
    assert math.isclose(result.jy, 2)
    assert math.isclose(result.jz, 2)


def test_cylindrical_tank_volume_matches_pi_r2h():
    tank = {"贮箱ID": "T-1", "推进剂密度rho_phi": 1.0, "总高H": 10.0}
    p1 = {
        "R11": 1, "R12": 0, "R13": 0, "R14": 0, "R15": 0, "R16": 0,
        "H11": 10, "H12": 10, "H13": 10, "H14": 10, "H15": 10, "H16": 10, "H17": 10, "H18": 10, "H19": 10,
        "h11": 0, "h12": 0, "semi_a11": 0, "b11": 1, "b12": 1, "beta11_deg": 0, "beta12_deg": 0,
    }
    p2 = {"是否启用扣除": "否"}
    res = propellant_by_level(tank, p1, p2, pd.DataFrame(), 201)
    full = res.iloc[-1]
    assert np.isclose(full["V_phi"], math.pi * 10, rtol=1e-4)
    assert np.isclose(full["X_phi"], 5, rtol=1e-4)


def test_equivalent_cylinder_without_qj_type1_geometry():
    tank = {"贮箱ID": "T-2", "几何类型": "等效圆柱", "等效半径R": 2.0, "等效高度H": 7.0, "推进剂密度rho_phi": 3.0}
    res = propellant_by_level(tank, {}, {}, pd.DataFrame(), 201)
    full = res.iloc[-1]
    assert np.isclose(full["V_phi"], math.pi * 2.0**2 * 7.0, rtol=1e-4)
    assert np.isclose(full["M_phi"], math.pi * 2.0**2 * 7.0 * 3.0, rtol=1e-4)
    assert np.isclose(full["X_phi"], 3.5, rtol=1e-4)


def test_invariant_mass_drops_after_separation_time():
    inv = pd.DataFrame([
        {"类型": "子级", "编号": 1, "M": 100, "X": 0, "Y": 0, "Z": 0, "Jx": 0, "Jy": 0, "Jz": 0},
        {"类型": "子级", "编号": 2, "M": 10, "X": 10, "Y": 0, "Z": 0, "Jx": 0, "Jy": 0, "Jz": 0},
    ])
    events = pd.DataFrame([
        {"类型": "子级", "编号": 1, "分离时间ts": 10},
        {"类型": "子级", "编号": 2, "分离时间ts": 20},
    ])
    assert invariant_mp_at_time(0, inv, events).mass == 110
    assert invariant_mp_at_time(10, inv, events).mass == 110
    assert invariant_mp_at_time(10.1, inv, events).mass == 10
    assert invariant_mp_at_time(20, inv, events).mass == 10
    assert invariant_mp_at_time(20.1, inv, events).mass == 0


def test_tank_uses_mef_during_burn_and_keeps_at_separation_time():
    level = pd.DataFrame([
        {"M_v": 0, "X_v": 0, "Y_v": 0, "Z_v": 0, "Jx_v": 0, "Jy_v": 0, "Jz_v": 0},
        {"M_v": 100, "X_v": 1, "Y_v": 0, "Z_v": 0, "Jx_v": 0, "Jy_v": 10, "Jz_v": 10},
        {"M_v": 200, "X_v": 2, "Y_v": 0, "Z_v": 0, "Jx_v": 0, "Jy_v": 20, "Jz_v": 20},
    ])
    tank = pd.Series({"类型": "子级", "秒流量mdot": 10, "加注量Mf": 200, "启动后剩余量Mef": 150, "关机剩余量Mr": 50, "坐标原点X或Loki": 10, "点火时间tf": 0, "关机时间tb": 10, "分离时间ts": 20})
    assert tank_mp_at_time(5, tank, level).mass == 100
    assert tank_mp_at_time(20, tank, level).mass == 50
    assert tank_mp_at_time(20.1, tank, level).mass == 0


def test_tank_uses_explicit_fill_and_shutdown_state_when_available():
    level = pd.DataFrame([
        {"M_v": 100, "X_v": 1, "Y_v": 0, "Z_v": 0, "Jx_v": 3, "Jy_v": 10, "Jz_v": 10},
        {"M_v": 200, "X_v": 2, "Y_v": 0, "Z_v": 0, "Jx_v": 4, "Jy_v": 20, "Jz_v": 20},
    ])
    tank = pd.Series({"类型": "子级", "秒流量mdot": 10, "加注量Mf": 200, "启动后剩余量Mef": 200, "关机剩余量Mr": 100, "加注后Xf": 5, "关机后Xr": 6, "加注后Jyf": 50, "关机后Jyr": 60, "坐标原点X或Loki": 10, "点火时间tf": 1, "关机时间tb": 2, "分离时间ts": 3})
    before = tank_mp_at_time(0, tank, level)
    after = tank_mp_at_time(2.5, tank, level)
    assert before.mass == 200
    assert before.x == 5
    assert before.jy == 50
    assert after.mass == 100
    assert after.x == 4
    assert after.jy == 60


def test_transport_results_combines_sections():
    data = {"运输状态": pd.DataFrame([
        {"部段编号": 1, "部段名称": "A", "Xq": 10, "M": 1, "Xcq": 0, "Ycq": 0, "Zcq": 0, "Jxcq": 0, "Jycq": 0, "Jzcq": 0},
        {"部段编号": 1, "部段名称": "A", "Xq": 10, "M": 1, "Xcq": 2, "Ycq": 0, "Zcq": 0, "Jxcq": 0, "Jycq": 0, "Jzcq": 0},
    ])}
    res = transport_results(data)
    assert math.isclose(res.iloc[0]["M_cq"], 2)
    assert math.isclose(res.iloc[0]["X_cq"], 1)
    assert math.isclose(res.iloc[0]["Jy_cq"], 2)


def test_bottom_level_results_from_hb():
    level = pd.DataFrame([
        {"贮箱ID": "子级-1-1", "h": 0, "V_phi": 0, "M_phi": 0, "X_phi": 0, "Jy_phi": 0},
        {"贮箱ID": "子级-1-1", "h": 1, "V_phi": 1, "M_phi": 1, "X_phi": 0.5, "Jy_phi": 1},
        {"贮箱ID": "子级-1-1", "h": 2, "V_phi": 2, "M_phi": 2, "X_phi": 1.0, "Jy_phi": 2},
    ])
    data = {"贮箱基本参数": pd.DataFrame([{"类型": "子级", "所属编号": 1, "贮箱编号": 1, "箱底hb": 1, "秒流量mdot": 1, "加注量Mf": 10, "启动后剩余量Mef": 10}])}
    res = build_tank_bottom_level_results(data, level)
    last = res.iloc[-1]
    assert math.isclose(last["h_b"], 1)
    assert math.isclose(last["V_b"], 1)
    assert math.isclose(last["M_b"], 1)
    assert math.isclose(last["T_b"], 8)
