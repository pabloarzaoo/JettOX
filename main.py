from __future__ import annotations
import os, sys, json, subprocess, shutil, csv, time, warnings
from datetime import datetime, UTC
from typing import List, Dict, Any

warnings.filterwarnings("ignore", category=DeprecationWarning)

try:
    import psutil
except:
    psutil = None

try:
    from rich.console import Console
    from rich.table import Table
    from rich.panel import Panel
    from rich.progress import Progress, SpinnerColumn, TextColumn, TimeElapsedColumn
    from rich.align import Align
    from rich.text import Text
    from rich import box
except:
    print("Install rich")
    raise

try:
    import pyfiglet
except:
    pyfiglet = None

ROOT = os.path.abspath(os.path.dirname(__file__))
FOLDERS = ["firmware", "hardware", "peripherals", "storage", "system"]
RESULT_DIR = os.path.join(ROOT, "result")
LOGS_DIR = os.path.join(ROOT, "logs")

COMMANDS = {
    "systeminfo": "systeminfo",
    "wmic_cpu": "wmic cpu get /format:list",
    "wmic_bios": "wmic bios get /format:list",
    "wmic_baseboard": "wmic baseboard get /format:list",
    "wmic_diskdrive": "wmic diskdrive get /format:list",
    "wmic_video": "wmic path win32_VideoController get /format:list",
    "driverquery": "driverquery /v /fo list",
    "tasklist_csv": "tasklist /V /FO CSV",
    "get_pnp_device": 'powershell -Command "Get-PnpDevice -PresentOnly | Format-List -Force"',
    "get_physicaldisk": 'powershell -Command "Get-PhysicalDisk | Format-List *"',
    "get_partition": 'powershell -Command "Get-Partition | Format-List *"',
    "get_netadapter": 'powershell -Command "Get-NetAdapter | Format-List *"',
    "wevtutil_system": "wevtutil qe System /c:200 /f:text",
    "wevtutil_app": "wevtutil qe Application /c:200 /f:text",
    "dxdiag": "dxdiag /t dxdiag_output.txt",
    "powercfg_query": "powercfg /q",
    "bcdedit": "bcdedit /enum all",
    "wmic_logicaldevice": 'wmic path CIM_LogicalDevice get /format:list',
    "wmic_service": "wmic service get /format:list",
}

console = Console()

class Logger:
    def __init__(self):
        self.lines = []

    def _append(self, level, msg):
        ts = datetime.now(UTC).isoformat() + "Z"
        self.lines.append(f"[{level}] {ts} {msg}")

    def info(self, msg): self._append("INFO", msg)
    def warn(self, msg): self._append("WARN", msg)
    def error(self, msg): self._append("ERROR", msg)

    def dump_to_file(self, path):
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(self.lines))
            return True
        except Exception as e:
            return False

logger = Logger()

def run(cmd, timeout=60):
    logger.info(f"Running: {cmd}")
    try:
        proc = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=timeout)
        return {"rc": proc.returncode, "out": proc.stdout or "", "err": proc.stderr or ""}
    except subprocess.TimeoutExpired as e:
        logger.warn(f"Timeout: {cmd}")
        return {"rc": -1, "out": e.stdout or "", "err": f"TIMEOUT: {e}"}
    except Exception as e:
        logger.error(f"Error: {cmd} - {e}")
        return {"rc": -1, "out": "", "err": str(e)}

def ensure_dirs():
    for d in FOLDERS: os.makedirs(os.path.join(ROOT, d), exist_ok=True)
    os.makedirs(RESULT_DIR, exist_ok=True)
    os.makedirs(LOGS_DIR, exist_ok=True)

def write_summary_and_raw(folder, summary, raw_blocks):
    p = os.path.join(ROOT, folder)
    os.makedirs(p, exist_ok=True)
    with open(os.path.join(p, "summary.json"), "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)
    with open(os.path.join(p, "raw.txt"), "w", encoding="utf-8", errors="replace") as f:
        f.write("==== RAW OUTPUT ====\n")
        for block in raw_blocks:
            f.write(block + ("\n" if not block.endswith("\n") else "") + "\n----\n\n")

def parse_wmic(text):
    d = {}
    for line in text.splitlines():
        if "=" in line: k, v = line.split("=", 1); d[k.strip()] = v.strip()
    return d

def safe_read_file(path):
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f: return f.read()
    except: return f"<error reading {path}>"

def collect_system():
    start = time.perf_counter()
    raw_blocks, summary = [], {"collected_at": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")}
    r = run(COMMANDS["systeminfo"], 30)
    raw_blocks.append("### systeminfo\n" + r["out"] + "\nERR:\n" + r["err"])
    summary["systeminfo_parsed"] = {k.strip(): v.strip() for line in r["out"].splitlines() if ":" in line for k, v in [line.split(":", 1)]}
    r2 = run(COMMANDS["tasklist_csv"], 30)
    raw_blocks.append("### tasklist\n" + r2["out"] + "\nERR:\n" + r2["err"])
    try:
        rows = list(csv.reader(r2["out"].splitlines()))
        summary["process_sample"] = [dict(zip(rows[0], row)) for row in rows[1:41]] if rows else []
    except: summary["process_sample_parse_error"] = "CSV error"
    for cmd in ["wevtutil_system", "wevtutil_app", "driverquery", "wmic_service"]:
        r = run(COMMANDS[cmd], 40)
        raw_blocks.append(f"### {cmd}\n" + r["out"] + "\nERR:\n" + r["err"])
    summary["drivers_count"] = r["out"].count("Driver Name") if "driverquery" in COMMANDS else 0
    summary["services_length"] = len(r["out"]) if "wmic_service" in COMMANDS else 0
    summary["elapsed_seconds"] = time.perf_counter() - start
    write_summary_and_raw("system", summary, raw_blocks)
    logger.info(f"System done in {summary['elapsed_seconds']:.2f}s")
    return summary

def collect_hardware():
    start = time.perf_counter()
    raw_blocks, summary = [], {"collected_at": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")}
    for cmd in ["wmic_cpu", "wmic_bios", "wmic_baseboard", "wmic_video"]:
        r = run(COMMANDS[cmd], 20)
        raw_blocks.append(f"### {cmd}\n" + r["out"] + "\nERR:\n" + r["err"])
        summary[cmd.split("_")[1]] = parse_wmic(r["out"])
    summary["video_raw_length"] = len(r["out"])
    if psutil:
        try:
            summary["logical_cpus"] = psutil.cpu_count(True)
            summary["physical_cpus"] = psutil.cpu_count(False)
            cf = psutil.cpu_freq()
            if cf: summary["cpu_freq_mhz"] = {"current": cf.current, "min": cf.min, "max": cf.max}
            vm = psutil.virtual_memory()
            summary["memory_total_bytes"] = vm.total
            summary["memory_available_bytes"] = vm.available
        except: summary["psutil_error"] = "Error"
    else: raw_blocks.append("### psutil not available\n")
    summary["elapsed_seconds"] = time.perf_counter() - start
    write_summary_and_raw("hardware", summary, raw_blocks)
    logger.info(f"Hardware done in {summary['elapsed_seconds']:.2f}s")
    return summary

def collect_firmware():
    start = time.perf_counter()
    raw_blocks, summary = [], {"collected_at": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")}
    r = run(COMMANDS["wmic_bios"], 15)
    raw_blocks.append("### wmic bios\n" + r["out"] + "\nERR:\n" + r["err"])
    for cmd in ["bcdedit", "powercfg_query"]:
        r = run(COMMANDS[cmd], 20)
        raw_blocks.append(f"### {cmd}\n" + r["out"] + "\nERR:\n" + r["err"])
    dx_out = os.path.join(ROOT, "dxdiag_output.txt")
    if os.path.exists(dx_out): os.remove(dx_out)
    r = run(COMMANDS["dxdiag"], 30)
    raw_blocks.append("### dxdiag\n" + (safe_read_file(dx_out) if os.path.exists(dx_out) else r["out"] + "\nERR:\n" + r["err"]))
    r = run('powershell -Command "Get-WinEvent -FilterHashtable @{LogName=\'System\';Level=3} -MaxEvents 200 | Where-Object { $_.Message -match \'microcode\' } | Format-List -Property TimeCreated,Id,Message"', 30)
    raw_blocks.append("### microcode events\n" + r["out"] + "\nERR:\n" + r["err"])
    parsed = parse_wmic(r["out"])
    for k in ["Manufacturer", "SMBIOSBIOSVersion", "BIOSVersion", "SerialNumber", "ReleaseDate"]:
        if k in parsed: summary[k] = parsed[k]
    summary["elapsed_seconds"] = time.perf_counter() - start
    write_summary_and_raw("firmware", summary, raw_blocks)
    logger.info(f"Firmware done in {summary['elapsed_seconds']:.2f}s")
    return summary

def collect_storage():
    start = time.perf_counter()
    raw_blocks, summary = [], {"collected_at": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")}
    for cmd in ["wmic_diskdrive", "get_physicaldisk", "get_partition"]:
        r = run(COMMANDS[cmd], 30)
        raw_blocks.append(f"### {cmd}\n" + r["out"] + "\nERR:\n" + r["err"])
    if shutil.which("smartctl"):
        r = run("smartctl --scan", 20)
        raw_blocks.append("### smartctl scan\n" + r["out"] + "\nERR:\n" + r["err"])
        for line in r["out"].splitlines():
            dev = line.split()[0]
            if dev:
                r = run(f"smartctl -H {dev}", 20)
                raw_blocks.append(f"### smartctl {dev}\n" + r["out"] + "\nERR:\n" + r["err"])
    else: raw_blocks.append("### smartctl not found\n")
    if psutil:
        try:
            summary["partitions_sample"] = [{"device": p.device, "mountpoint": p.mountpoint, "fstype": p.fstype, **psutil.disk_usage(p.mountpoint)._asdict()} for p in psutil.disk_partitions(True)]
        except: summary["partitions_error"] = "Error"
    summary["elapsed_seconds"] = time.perf_counter() - start
    write_summary_and_raw("storage", summary, raw_blocks)
    logger.info(f"Storage done in {summary['elapsed_seconds']:.2f}s")
    return summary

def collect_peripherals():
    start = time.perf_counter()
    raw_blocks, summary = [], {"collected_at": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")}
    usb = run('powershell -Command "Get-PnpDevice -PresentOnly | Where-Object { $_.InstanceId -like \'USB*\' } | Format-List -Property *"', 40)
    raw_blocks.append("### USB\n" + usb["out"] + "\nERR:\n" + usb["err"])
    for cmd in ["get_pnp_device", "get_netadapter", "wmic_logicaldevice"]:
        r = run(COMMANDS[cmd], 40 if "pnp" in cmd else 30)
        raw_blocks.append(f"### {cmd}\n" + r["out"] + "\nERR:\n" + r["err"])
    bt = run('powershell -Command "Get-PnpDevice -PresentOnly | Where-Object { $_.Class -like \'Bluetooth\' } | Format-List -Property *"', 20)
    raw_blocks.append("### Bluetooth\n" + bt["out"] + "\nERR:\n" + bt["err"])
    summary["usb_entries_length"] = len(usb["out"])
    summary["pnp_entries_length"] = len(r["out"]) if "pnp" in cmd else 0
    summary["netadapter_entries_length"] = len(r["out"]) if "netadapter" in cmd else 0
    summary["elapsed_seconds"] = time.perf_counter() - start
    write_summary_and_raw("peripherals", summary, raw_blocks)
    logger.info(f"Peripherals done in {summary['elapsed_seconds']:.2f}s")
    return summary

def aggregate_results(summary_map):
    with open(os.path.join(RESULT_DIR, "results_of_all_files.json"), "w", encoding="utf-8") as f:
        json.dump(summary_map, f, indent=2, ensure_ascii=False)
    logger.info("Aggregated results written")

def is_admin():
    try: import ctypes; return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except: return False

def render_header():
    if pyfiglet:
        banner = pyfiglet.figlet_format("JettOX Pro", font="slant")
        console.print(f"[bold green]{banner}[/bold green]")
    else:
        console.rule("[bold green]JettOX Pro - Deep Diagnostics[/bold green]")

def run_module_with_spinner(name, fn, status_map):
    with Progress(SpinnerColumn(style="green"), TextColumn("[progress.description]{task.description}"), TimeElapsedColumn()) as progress:
        task = progress.add_task(f"[bold bright_magenta]{name}[/bold bright_magenta] — collecting...", start=False)
        progress.start_task(task)
        try:
            res = fn()
            status_map[name] = True
        except Exception as e:
            res = {"error": str(e)}
            status_map[name] = False
        progress.stop()
    return res

def main():
    ensure_dirs()
    started = datetime.now()
    logname = datetime.now().strftime("run_%Y-%m-%d_%H%M%S.txt")
    logfile = os.path.join(LOGS_DIR, logname)
    console.clear()
    render_header()
    admin = is_admin()
    logger.info(f"Admin: {admin}")
    if not admin: logger.warn("Not admin - limited queries")
    summary_map = {"collected_at": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"), "local_start": started.isoformat(), "admin": admin}
    modules = [("system", collect_system), ("hardware", collect_hardware), ("firmware", collect_firmware), ("storage", collect_storage), ("peripherals", collect_peripherals)]
    live_status = {name: False for name, _ in modules}
    console.print("\n[bold white]Starting collection...[/bold white]\n")
    for name, fn in modules:
        console.print(Panel(f"[bold]{name.upper()}[/bold] — starting", subtitle="Read-only", box=box.SIMPLE))
        res = run_module_with_spinner(name, fn, live_status)
        summary_map[name] = res
        elapsed = res.get("elapsed_seconds", 0)
        console.print(f"[green]✔[/green] {name} finished in {elapsed:.2f}s\n")
    total_elapsed = sum(summary_map.get(m, {}).get("elapsed_seconds", 0) for m, _ in modules)
    summary_map["total_elapsed_seconds"] = total_elapsed
    aggregate_results(summary_map)
    logger.dump_to_file(logfile)
    console.rule("[bold green]Run Complete[/bold green]")
    console.print(Panel(Align.center(Text(f"Check folders in: {ROOT}\nResults: {RESULT_DIR}\nLogs: {LOGS_DIR}\nTotal: {total_elapsed:.2f}s", justify="center")), title="[bold magenta]JettOX Pro[/bold magenta]", subtitle="Read-Only Diagnostics", box=box.DOUBLE))

if __name__ == "__main__":
    main()