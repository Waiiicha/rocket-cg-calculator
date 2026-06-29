import math

import numpy as np
import pandas as pd

from rocket_mass_properties import MassProperties, combine_mass_properties, invariant_mp_at_time, propellant_by_level


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


def test_invariant_mass_drops_at_separation_time():
    inv = pd.DataFrame([
        {"类型": "子级", "编号": 1, "M": 100, "X": 0, "Y": 0, "Z": 0, "Jx": 0, "Jy": 0, "Jz": 0},
        {"类型": "子级", "编号": 2, "M": 10, "X": 10, "Y": 0, "Z": 0, "Jx": 0, "Jy": 0, "Jz": 0},
    ])
    events = pd.DataFrame([
        {"类型": "子级", "编号": 1, "分离时间ts": 10},
        {"类型": "子级", "编号": 2, "分离时间ts": 20},
    ])
    assert invariant_mp_at_time(0, inv, events).mass == 110
    assert invariant_mp_at_time(10, inv, events).mass == 10
    assert invariant_mp_at_time(20, inv, events).mass == 0
