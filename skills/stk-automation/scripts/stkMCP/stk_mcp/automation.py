"""Create five STK 11 UAV external-ephemeris link cases and export STK data.

Run:
    python create_uav_stk11_project.py

The exported CSV fields are read from STK Data Providers, not recomputed here.
"""

from __future__ import annotations

import csv
import copy
import json
import math
import statistics
from datetime import datetime, timedelta, timezone
from pathlib import Path

try:
    import win32com.client
    import win32com.client.dynamic
except ImportError:
    win32com = None


CONFIG = {
    "scenario_name": "UAV_B1_L1_5Groups_ExternalEphem",
    "output_dir": "stk_uav_outputs",
    "show_stk": True,
    "overwrite_current_scenario": True,
    "start_utc": "2026-05-09T04:00:00Z",
    "sample_step_s": 1.0,
    "frequencies_hz": {
        "B1": 1_561_098_000.0,
        "L1": 1_575_420_000.0,
    },
    "groups": [
        {
            "id": "G01",
            "trajectory_type": "baseline_arc",
            "uav_name": "UAV_G01",
            "ground_station_name": "Radiator_G01",
            "altitude_m": 7500.0,
            "start_speed_kmh": 700.0,
            "stop_speed_kmh": 800.0,
            "start_lat_deg": 39.8300,
            "start_lon_deg": 116.2200,
            "stop_lat_deg": 39.9700,
            "stop_lon_deg": 116.6200,
            "cross_track_m": 900.0,
            "radiator_lat_deg": 39.9000,
            "radiator_lon_deg": 116.3900,
            "radiator_altitude_m": 50.0,
        },
        {
            "id": "G02",
            "trajectory_type": "s_curve",
            "uav_name": "UAV_G02",
            "ground_station_name": "Radiator_G02",
            "altitude_m": 7200.0,
            "start_speed_kmh": 720.0,
            "stop_speed_kmh": 790.0,
            "start_lat_deg": 31.0900,
            "start_lon_deg": 121.1800,
            "stop_lat_deg": 31.3900,
            "stop_lon_deg": 121.7800,
            "cross_track_m": 1800.0,
            "radiator_lat_deg": 31.2300,
            "radiator_lon_deg": 121.4700,
            "radiator_altitude_m": 30.0,
        },
        {
            "id": "G03",
            "trajectory_type": "smooth_dogleg",
            "uav_name": "UAV_G03",
            "ground_station_name": "Radiator_G03",
            "altitude_m": 7800.0,
            "start_speed_kmh": 700.0,
            "stop_speed_kmh": 780.0,
            "start_lat_deg": 34.1000,
            "start_lon_deg": 108.6500,
            "stop_lat_deg": 34.4300,
            "stop_lon_deg": 109.2800,
            "cross_track_m": 2200.0,
            "radiator_lat_deg": 34.2500,
            "radiator_lon_deg": 108.9500,
            "radiator_altitude_m": 450.0,
        },
        {
            "id": "G04",
            "trajectory_type": "partial_orbit",
            "uav_name": "UAV_G04",
            "ground_station_name": "Radiator_G04",
            "altitude_m": 7600.0,
            "start_speed_kmh": 710.0,
            "stop_speed_kmh": 800.0,
            "start_lat_deg": 30.4000,
            "start_lon_deg": 114.0000,
            "stop_lat_deg": 30.7800,
            "stop_lon_deg": 114.6400,
            "cross_track_m": 2600.0,
            "radiator_lat_deg": 30.5900,
            "radiator_lon_deg": 114.3000,
            "radiator_altitude_m": 60.0,
        },
        {
            "id": "G05",
            "trajectory_type": "three_d_maneuver",
            "uav_name": "UAV_G05",
            "ground_station_name": "Radiator_G05",
            "altitude_m": 7400.0,
            "altitude_profile_m": [7300.0, 7800.0, 7400.0],
            "start_speed_kmh": 700.0,
            "stop_speed_kmh": 785.0,
            "start_lat_deg": 22.9600,
            "start_lon_deg": 112.9500,
            "stop_lat_deg": 23.3200,
            "stop_lon_deg": 113.5800,
            "cross_track_m": 1800.0,
            "radiator_lat_deg": 23.1300,
            "radiator_lon_deg": 113.2600,
            "radiator_altitude_m": 25.0,
        },
    ],
}


WGS84_A = 6_378_137.0
WGS84_F = 1.0 / 298.257223563
WGS84_E2 = WGS84_F * (2.0 - WGS84_F)


def default_config() -> dict:
    return copy.deepcopy(CONFIG)


def load_config(path: str | Path) -> dict:
    with Path(path).open("r", encoding="utf-8") as handle:
        return json.load(handle)


def run_automation(config: dict | None = None, base_dir: str | Path | None = None) -> dict:
    global CONFIG
    CONFIG = copy.deepcopy(config) if config is not None else default_config()
    base_path = Path(base_dir).resolve() if base_dir is not None else Path.cwd().resolve()
    out_dir = base_path / CONFIG["output_dir"]
    ephem_dir = out_dir / "external_ephemeris"
    out_dir.mkdir(parents=True, exist_ok=True)
    ephem_dir.mkdir(parents=True, exist_ok=True)

    start = parse_utc(CONFIG["start_utc"])
    group_tracks = []
    for group in CONFIG["groups"]:
        track = build_track(group, start)
        ephem_path = ephem_dir / f"{group['id']}_{group['uav_name']}.e"
        write_ephemeris(ephem_path, start, track)
        group_tracks.append({"config": group, "track": track, "ephem_path": ephem_path})

    stop = max(item["track"][-1]["time"] for item in group_tracks) + timedelta(seconds=5.0)
    scenario_path = out_dir / f"{CONFIG['scenario_name']}.sc"

    app, root = connect_stk()
    create_stk_scenario(root, scenario_path, start, stop, group_tracks)
    exported = export_all_tables(root, out_dir, start, stop, group_tracks)

    summary = validate_exports(exported, CONFIG["groups"])
    result = {
        "scenario_path": str(scenario_path),
        "output_dir": str(out_dir),
        "ephemeris_paths": {
            item["config"]["id"]: str(item["ephem_path"]) for item in group_tracks
        },
        "exported_paths": {key: str(path) for key, path in exported.items()},
        "validation": summary,
    }
    if CONFIG["show_stk"]:
        app.Visible = True
    return result


def main() -> None:
    result = run_automation(CONFIG, Path(__file__).resolve().parent.parent)
    scenario_path = result["scenario_path"]
    exported = result["exported_paths"]
    group_tracks = [
        {"config": group, "ephem_path": result["ephemeris_paths"][group["id"]]}
        for group in CONFIG["groups"]
    ]
    print(f"STK scenario saved: {scenario_path}")
    for item in group_tracks:
        group = item["config"]
        print(
            f"{group['id']}: altitude={group['altitude_m']:.0f} m, "
            f"speed={group['start_speed_kmh']:.1f}->{group['stop_speed_kmh']:.1f} km/h, "
            f"ephemeris={item['ephem_path']}"
        )
    for label, path in exported.items():
        print(f"{label} exported: {path}")


def connect_stk():
    if win32com is None:
        raise RuntimeError("pywin32 is required on Windows with STK installed: pip install pywin32")
    try:
        app = win32com.client.Dispatch("STK11.Application")
    except Exception:
        app = win32com.client.dynamic.Dispatch("STK11.Application")
    app.Visible = bool(CONFIG["show_stk"])
    return app, app.Personality2


def create_stk_scenario(root, scenario_path: Path, start: datetime, stop: datetime, group_tracks: list[dict]) -> None:
    if CONFIG["overwrite_current_scenario"]:
        try:
            root.CloseScenario()
        except Exception:
            pass

    root.NewScenario(CONFIG["scenario_name"])
    scenario = root.CurrentScenario
    set_units(root)
    scenario.SetTimePeriod(stk_time(start), stk_time(stop))
    root.Rewind()

    for item in group_tracks:
        group = item["config"]
        facility = scenario.Children.New(8, group["ground_station_name"])
        facility.Position.AssignGeodetic(
            float(group["radiator_lat_deg"]),
            float(group["radiator_lon_deg"]),
            float(group["radiator_altitude_m"]),
        )

        aircraft = scenario.Children.New(1, group["uav_name"])
        aircraft.SetRouteType(6)
        route = aircraft.Route
        route.Filename = str(item["ephem_path"].resolve())
        route.Propagate()

        receiver = aircraft.Children.New(17, "UAV_Rx")
        receiver.SetModel("Simple Receiver Model")
        try:
            receiver.Model.AutoTrackFrequency = True
            receiver.Model.AutoScaleBandwidth = True
        except Exception:
            pass

        for band, freq_hz in CONFIG["frequencies_hz"].items():
            transmitter = facility.Children.New(24, f"{group['id']}_{band}_Tx")
            transmitter.SetModel("Simple Transmitter Model")
            transmitter.Model.Frequency = float(freq_hz)

            chain = scenario.Children.New(4, f"{group['id']}_{band}_Link")
            chain.Objects.Add(transmitter.Path)
            chain.Objects.Add(receiver.Path)
            chain.ComputeAccess()

            access = transmitter.GetAccessToObject(receiver)
            access.ComputeAccess()
            add_doppler_scalar(access, group["id"], band)

    save_scenario(root, scenario_path)


def export_all_tables(root, out_dir: Path, start: datetime, stop: datetime, group_tracks: list[dict]) -> dict[str, Path]:
    set_units(root)
    start_s = stk_time(start)
    stop_s = stk_time(stop)
    step = float(CONFIG["sample_step_s"])
    fieldnames = [
        "Group",
        "TimeUTC",
        "Band",
        "X_m",
        "Y_m",
        "Z_m",
        "Vx_mps",
        "Vy_mps",
        "Vz_mps",
        "Ax_mps2",
        "Ay_mps2",
        "Az_mps2",
        "Doppler_Hz",
        "DopplerRate_Hzps",
    ]

    all_rows: list[dict[str, str]] = []
    exported: dict[str, Path] = {}
    for item in group_tracks:
        group = item["config"]
        group_id = group["id"]
        aircraft = root.GetObjectFromPath(f"Aircraft/{group['uav_name']}")
        position = exec_provider(aircraft.DataProviders.Item("Cartesian Position").Group.Item("Fixed"), start_s, stop_s, step)
        velocity = exec_provider(aircraft.DataProviders.Item("Cartesian Velocity").Group.Item("Fixed"), start_s, stop_s, step)
        acceleration = exec_provider(
            aircraft.DataProviders.Item("Cartesian Acceleration").Group.Item("Fixed"), start_s, stop_s, step
        )
        state_rows = state_rows_by_time(group_id, position, velocity, acceleration)
        group_rows: list[dict[str, str]] = []

        for band in CONFIG["frequencies_hz"]:
            tx = root.GetObjectFromPath(f"Facility/{group['ground_station_name']}/Transmitter/{group_id}_{band}_Tx")
            rx = root.GetObjectFromPath(f"Aircraft/{group['uav_name']}/Receiver/UAV_Rx")
            access = tx.GetAccessToObject(rx)
            link = exec_provider(access.DataProviders.Item("Link Information"), start_s, stop_s, step)
            rate = exec_provider(
                access.DataProviders.Item("Scalar Calculations").Group.Item(f"{group_id}_{band}_DopplerShift"),
                start_s,
                stop_s,
                step,
            )

            band_rows = merge_link_rows(group_id, band, state_rows, link, rate)
            group_rows.extend(band_rows)
            all_rows.extend(band_rows)

            band_path = out_dir / f"{group_id}_{group['trajectory_type']}_{band}_uav_receiver_link.csv"
            write_csv(band_path, fieldnames, band_rows)
            exported[f"{group_id}_{band}"] = band_path

        group_path = out_dir / f"{group_id}_{group['trajectory_type']}_all_frequencies_uav_receiver_link.csv"
        write_csv(group_path, fieldnames, group_rows)
        exported[f"{group_id}_ALL"] = group_path

    all_path = out_dir / "five_groups_uav_link_doppler_all.csv"
    write_csv(all_path, fieldnames, all_rows)
    exported["ALL_GROUPS"] = all_path
    return exported


def state_rows_by_time(group_id: str, position: dict[str, list], velocity: dict[str, list], acceleration: dict[str, list]):
    rows = {}
    for idx, time_value in enumerate(position["Time"]):
        key = normalize_stk_time(time_value)
        rows[key] = {
            "Group": group_id,
            "TimeUTC": key,
            "X_m": fmt(position["x"][idx]),
            "Y_m": fmt(position["y"][idx]),
            "Z_m": fmt(position["z"][idx]),
            "Vx_mps": fmt(velocity["x"][idx]),
            "Vy_mps": fmt(velocity["y"][idx]),
            "Vz_mps": fmt(velocity["z"][idx]),
            "Ax_mps2": fmt(acceleration["x"][idx]),
            "Ay_mps2": fmt(acceleration["y"][idx]),
            "Az_mps2": fmt(acceleration["z"][idx]),
        }
    return rows


def merge_link_rows(group_id: str, band: str, state_rows: dict[str, dict], link: dict[str, list], rate: dict[str, list]):
    rows = []
    for idx, time_value in enumerate(link["Time"]):
        key = normalize_stk_time(time_value)
        if key not in state_rows:
            continue
        row = dict(state_rows[key])
        row["Group"] = group_id
        row["Band"] = band
        row["Doppler_Hz"] = fmt(link["Freq. Doppler Shift"][idx])
        row["DopplerRate_Hzps"] = fmt(rate["Scalar Rate"][idx])
        rows.append(row)
    return rows


def add_doppler_scalar(access, group_id: str, band: str) -> None:
    scalars = access.Vgt.CalcScalars
    name = f"{group_id}_{band}_DopplerShift"
    try:
        if scalars.Contains(name):
            scalars.Remove(name)
    except Exception:
        pass
    scalars.Factory.CreateCalcScalarDataElement(
        name,
        f"{group_id} {band} Doppler shift from STK Link Information",
        "Link Information",
        "Freq. Doppler Shift",
    )


def exec_provider(provider, start_s: str, stop_s: str, step_s: float) -> dict[str, list]:
    result = provider.Exec(start_s, stop_s, step_s)
    data = {}
    for idx in range(result.DataSets.Count):
        dataset = result.DataSets.Item(idx)
        data[dataset.ElementName] = list(dataset.GetValues())
    return data


def build_track(group: dict, start: datetime) -> list[dict]:
    sample_s = float(CONFIG["sample_step_s"])
    origin_lat = 0.5 * (float(group["start_lat_deg"]) + float(group["stop_lat_deg"]))
    origin_lon = 0.5 * (float(group["start_lon_deg"]) + float(group["stop_lon_deg"]))
    origin_alt = profile_altitude(group, 0.5)
    origin = lla_to_ecef(origin_lat, origin_lon, origin_alt)
    east, north, up = enu_basis(origin_lat, origin_lon)

    start_enu = lla_to_enu(
        float(group["start_lat_deg"]), float(group["start_lon_deg"]), origin_lat, origin_lon
    )
    stop_enu = lla_to_enu(
        float(group["stop_lat_deg"]), float(group["stop_lon_deg"]), origin_lat, origin_lon
    )
    dx = stop_enu[0] - start_enu[0]
    dy = stop_enu[1] - start_enu[1]
    length = math.hypot(dx, dy)
    if length <= 0:
        raise SystemExit(f"{group['id']} start and stop positions must be different.")
    nx = -dy / length
    ny = dx / length

    dense_n = 5001
    dense_s = [i / (dense_n - 1) for i in range(dense_n)]
    dense_pos = [curve_position(group, s, start_enu, dx, dy, nx, ny, origin, east, north, up, origin_alt) for s in dense_s]
    arc_len = [0.0]
    for p0, p1 in zip(dense_pos[:-1], dense_pos[1:]):
        arc_len.append(arc_len[-1] + distance(p0, p1))

    v0 = kmh_to_mps(float(group["start_speed_kmh"]))
    v1 = kmh_to_mps(float(group["stop_speed_kmh"]))
    time_grid = [0.0]
    for i in range(1, dense_n):
        ds = arc_len[i] - arc_len[i - 1]
        v_mid = 0.5 * (speed_at_s(dense_s[i - 1], v0, v1) + speed_at_s(dense_s[i], v0, v1))
        time_grid.append(time_grid[-1] + ds / v_mid)

    duration_s = time_grid[-1]
    n_samples = int(math.ceil(duration_s / sample_s)) + 1
    samples = []
    for i in range(n_samples):
        time_s = min(i * sample_s, duration_s)
        s = interp_s_from_time(time_s, time_grid, dense_s)
        pos = curve_position(group, s, start_enu, dx, dy, nx, ny, origin, east, north, up, origin_alt)
        samples.append({"time": start + timedelta(seconds=time_s), "time_s": time_s, "pos": pos})

    attach_velocity(samples)
    return samples


def curve_position(group, s, start_enu, dx, dy, nx, ny, origin, east, north, up, origin_alt):
    cross, along_extra = trajectory_offsets(group, s)
    e = start_enu[0] + dx * s + nx * cross + (dx / max(math.hypot(dx, dy), 1.0)) * along_extra
    n_m = start_enu[1] + dy * s + ny * cross + (dy / max(math.hypot(dx, dy), 1.0)) * along_extra
    vertical = profile_altitude(group, s) - origin_alt
    return add_vec(origin, add_vec(scale_vec(east, e), add_vec(scale_vec(north, n_m), scale_vec(up, vertical))))


def trajectory_offsets(group: dict, s: float) -> tuple[float, float]:
    amp = float(group["cross_track_m"])
    kind = group.get("trajectory_type", "baseline_arc")
    if kind == "baseline_arc":
        return amp * math.sin(math.pi * s), 0.0
    if kind == "s_curve":
        return amp * math.sin(2.0 * math.pi * s), 0.0
    if kind == "smooth_dogleg":
        bend = math.tanh((s - 0.50) * 7.0) / math.tanh(3.5)
        return amp * 0.5 * bend, 0.0
    if kind == "partial_orbit":
        theta = math.radians(215.0 - 250.0 * s)
        radius = amp
        cross = radius * math.sin(theta)
        along_extra = radius * 0.45 * math.cos(theta)
        return cross, along_extra
    if kind == "three_d_maneuver":
        return amp * (0.65 * math.sin(math.pi * s) + 0.35 * math.sin(3.0 * math.pi * s)), 0.0
    raise SystemExit(f"Unsupported trajectory_type: {kind}")


def profile_altitude(group: dict, s: float) -> float:
    if group.get("trajectory_type") == "three_d_maneuver" and "altitude_profile_m" in group:
        low0, peak, low1 = [float(value) for value in group["altitude_profile_m"]]
        if s <= 0.5:
            q = s / 0.5
            return low0 + (peak - low0) * smooth01(q)
        q = (s - 0.5) / 0.5
        return peak + (low1 - peak) * smooth01(q)
    return float(group["altitude_m"])


def smooth01(t: float) -> float:
    t = max(0.0, min(1.0, t))
    return t * t * (3.0 - 2.0 * t)


def speed_at_s(s: float, v0: float, v1: float) -> float:
    return v0 + (v1 - v0) * s


def interp_s_from_time(time_s: float, time_grid: list[float], s_grid: list[float]) -> float:
    if time_s <= time_grid[0]:
        return s_grid[0]
    if time_s >= time_grid[-1]:
        return s_grid[-1]
    lo = 0
    hi = len(time_grid) - 1
    while hi - lo > 1:
        mid = (lo + hi) // 2
        if time_grid[mid] <= time_s:
            lo = mid
        else:
            hi = mid
    span = time_grid[hi] - time_grid[lo]
    frac = (time_s - time_grid[lo]) / span if span else 0.0
    return s_grid[lo] + (s_grid[hi] - s_grid[lo]) * frac


def attach_velocity(samples: list[dict]) -> None:
    positions = [sample["pos"] for sample in samples]
    times = [sample["time_s"] for sample in samples]
    for idx, sample in enumerate(samples):
        if idx == 0:
            j0, j1 = 0, 1
        elif idx == len(samples) - 1:
            j0, j1 = len(samples) - 2, len(samples) - 1
        else:
            j0, j1 = idx - 1, idx + 1
        dt = times[j1] - times[j0]
        sample["vel"] = tuple((positions[j1][axis] - positions[j0][axis]) / dt for axis in range(3))


def write_ephemeris(path: Path, epoch: datetime, samples: list[dict]) -> None:
    lines = [
        "stk.v.11.0",
        "BEGIN Ephemeris",
        f"NumberOfEphemerisPoints {len(samples)}",
        f"ScenarioEpoch {stk_time(epoch)}",
        "InterpolationMethod Lagrange",
        "InterpolationOrder 5",
        "CentralBody Earth",
        "CoordinateSystem Fixed",
        "DistanceUnit Meters",
        "EphemerisTimePosVel",
    ]
    for sample in samples:
        x, y, z = sample["pos"]
        vx, vy, vz = sample["vel"]
        lines.append(f"{sample['time_s']:.6f} {x:.9f} {y:.9f} {z:.9f} {vx:.12f} {vy:.12f} {vz:.12f}")
    lines.extend(["END Ephemeris", ""])
    path.write_text("\n".join(lines), encoding="ascii")


def set_units(root) -> None:
    unit_pairs = [
        ("DateFormat", "UTCG"),
        ("Distance", "m"),
        ("Time", "sec"),
        ("Latitude", "deg"),
        ("Longitude", "deg"),
        ("Angle", "deg"),
        ("Speed", "m/sec"),
        ("Acceleration", "m/sec^2"),
        ("Frequency", "Hz"),
    ]
    for dimension, unit in unit_pairs:
        try:
            root.UnitPreferences.SetCurrentUnit(dimension, unit)
        except Exception:
            pass


def save_scenario(root, scenario_path: Path) -> None:
    scenario_path = scenario_path.resolve()
    try:
        root.SaveScenarioAs(str(scenario_path))
    except Exception:
        root.ExecuteCommand(f'Save / "{scenario_path}"')


def parse_utc(value: str) -> datetime:
    if value.endswith("Z"):
        value = value[:-1] + "+00:00"
    dt = datetime.fromisoformat(value)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def stk_time(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).strftime("%d %b %Y %H:%M:%S.%f")[:-3]


def normalize_stk_time(value) -> str:
    text = str(value)
    if "." in text:
        prefix, frac = text.split(".", 1)
        return prefix + "." + frac[:3]
    return text


def lla_to_ecef(lat_deg: float, lon_deg: float, alt_m: float) -> tuple[float, float, float]:
    lat = math.radians(lat_deg)
    lon = math.radians(lon_deg)
    sin_lat = math.sin(lat)
    cos_lat = math.cos(lat)
    n = WGS84_A / math.sqrt(1.0 - WGS84_E2 * sin_lat * sin_lat)
    x = (n + alt_m) * cos_lat * math.cos(lon)
    y = (n + alt_m) * cos_lat * math.sin(lon)
    z = (n * (1.0 - WGS84_E2) + alt_m) * sin_lat
    return x, y, z


def enu_basis(lat_deg: float, lon_deg: float):
    lat = math.radians(lat_deg)
    lon = math.radians(lon_deg)
    east = (-math.sin(lon), math.cos(lon), 0.0)
    north = (-math.sin(lat) * math.cos(lon), -math.sin(lat) * math.sin(lon), math.cos(lat))
    up = (math.cos(lat) * math.cos(lon), math.cos(lat) * math.sin(lon), math.sin(lat))
    return east, north, up


def lla_to_enu(lat_deg: float, lon_deg: float, origin_lat_deg: float, origin_lon_deg: float):
    mean_lat = math.radians(origin_lat_deg)
    meters_per_deg_lat = 111_319.49079327358
    meters_per_deg_lon = meters_per_deg_lat * math.cos(mean_lat)
    east = (lon_deg - origin_lon_deg) * meters_per_deg_lon
    north = (lat_deg - origin_lat_deg) * meters_per_deg_lat
    return east, north


def add_vec(a, b):
    return a[0] + b[0], a[1] + b[1], a[2] + b[2]


def scale_vec(a, scale: float):
    return a[0] * scale, a[1] * scale, a[2] * scale


def distance(a, b):
    return math.sqrt((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2 + (a[2] - b[2]) ** 2)


def kmh_to_mps(value: float) -> float:
    return value / 3.6


def fmt(value) -> str:
    try:
        return f"{float(value):.9f}"
    except Exception:
        return str(value)


def write_csv(path: Path, fieldnames: list[str], rows: list[dict[str, str]]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def validate_exports(exported: dict[str, Path], groups: list[dict]) -> dict:
    validation = {}
    for group in groups:
        group_id = group["id"]
        trajectory = group.get("trajectory_type", "trajectory")
        b1_path = Path(exported[f"{group_id}_B1"])
        l1_path = Path(exported[f"{group_id}_L1"])
        b1_rows = read_csv_rows(b1_path)
        l1_rows = read_csv_rows(l1_path)
        speeds = [
            math.sqrt(sum(float(row[key]) ** 2 for key in ("Vx_mps", "Vy_mps", "Vz_mps"))) * 3.6
            for row in b1_rows
        ]
        accels = [
            math.sqrt(sum(float(row[key]) ** 2 for key in ("Ax_mps2", "Ay_mps2", "Az_mps2")))
            for row in b1_rows
        ]
        validation[group_id] = {
            "trajectory_type": trajectory,
            "rows_B1": len(b1_rows),
            "rows_L1": len(l1_rows),
            "times_match": [row["TimeUTC"] for row in b1_rows] == [row["TimeUTC"] for row in l1_rows],
            "bad_values": count_bad_values(b1_rows) + count_bad_values(l1_rows),
            "speed_kmh_min": min(speeds) if speeds else None,
            "speed_kmh_mean": statistics.mean(speeds) if speeds else None,
            "speed_kmh_max": max(speeds) if speeds else None,
            "accel_mps2_min": min(accels) if accels else None,
            "accel_mps2_mean": statistics.mean(accels) if accels else None,
            "accel_mps2_max": max(accels) if accels else None,
        }
    return validation


def read_csv_rows(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def count_bad_values(rows: list[dict[str, str]]) -> int:
    total = 0
    for row in rows:
        for value in row.values():
            if value == "" or value.lower() == "nan":
                total += 1
    return total


if __name__ == "__main__":
    main()
