import datetime
import fnmatch
import logging
import os
import subprocess
import time

import psutil
import win32api
import win32gui
import win32process

DEBUG = True
logging.basicConfig(level=logging.DEBUG if DEBUG else logging.INFO)
logger = logging.getLogger("main")

TARGET_PROCESS = "msedge.exe"
START_TIME = datetime.datetime.now()
TIMEOUT = 3 #Seconds

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
START_SCRIPT_PATH = os.path.join(ROOT_DIR, "open_browser.ps1")

WINDOW_TITLE_TO_MONITOR = {
    "TI * Dashboards - Grafana": 1,
    "NOC SCC: Dashboard": 0,
}

def get_process_name_from_hwnd(hwnd):
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        return pid, process.name()
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return None, "Desconhecido"

def move_window_to_monitor(hwnd, monitor_index, title):
    monitors = win32api.EnumDisplayMonitors()
    if monitor_index >= len(monitors):
        exc = Exception(f"Monitor {monitor_index} não existe")
        logger.exception(exc)
        raise exc

    mi = win32api.GetMonitorInfo(monitors[monitor_index][0])
    left, top, right, bottom = mi["Monitor"]
    width, height = right - left, bottom - top

    win32gui.MoveWindow(hwnd, left, top, width, height, True)
    win32gui.ShowWindow(hwnd, 9)
    win32gui.SetForegroundWindow(hwnd)
    logger.debug(f"Janela '{title}' movida para Monitor {monitor_index} ({left},{top})")

def enum_handler(hwnd, result):

    if win32gui.IsWindowVisible(hwnd):
        title = win32gui.GetWindowText(hwnd)
        if title:
            pid, pname = get_process_name_from_hwnd(hwnd)
            if pname.lower() == TARGET_PROCESS.lower():
                result.append((hwnd, pid, pname, title))

def title_matches_any_pattern(title: str, patterns: list[str]) -> bool:
    return any(fnmatch.fnmatchcase(title, pattern) for pattern in patterns)

def wait_windows(no_wait=False):

    def check():
        windows = []
        win32gui.EnumWindows(enum_handler, windows)
        
        matched_titles = {title for _, _, _, title in windows if title_matches_any_pattern(title, WINDOW_TITLE_TO_MONITOR.keys())}
        if all(any(fnmatch.fnmatchcase(title, pattern) for title in matched_titles) for pattern in WINDOW_TITLE_TO_MONITOR.keys()):
            logger.debug("Todas as janelas esperadas foram abertas.")
            logger.debug(f"Janelas abertas: {windows}")
            return windows
        else:
            logger.debug("Nem todas as janelas esperadas foram abertas.")
            logger.debug(f"Janelas abertas: {windows}")
            return None
        
    if no_wait:
        windows = check()
        
    while (datetime.datetime.now() - START_TIME).total_seconds() < TIMEOUT:
        windows = check()
        if windows:
            return windows
        time.sleep(0.1)
    logger.exception("Timeout: nem todas as janelas apareceram.")
    raise TimeoutError("Uma janela esperada não foi encontrada.")

def get_monitors():
    monitors = win32api.EnumDisplayMonitors()
    for i, monitor in enumerate(monitors):
        monitor_info = win32api.GetMonitorInfo(monitor[0])
        logger.debug(f"Monitor {i}: {monitor_info['Monitor']}")
        return monitors


def main():
    windows = wait_windows(no_wait=True)
    if not windows:
        logger.debug("abrindo browser...")
        subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", START_SCRIPT_PATH], check=True)
        windows = wait_windows()
    monitors = get_monitors()
    
    if len(windows) > len(monitors):
        exc = Exception("Mais janelas abertas do que monitores detectados.")
        logger.exception(exc)
        raise exc
    
    for hwnd, pid, pname, title in windows:
        for title_pattern, monitor_index in WINDOW_TITLE_TO_MONITOR.items():
            if fnmatch.fnmatchcase(title, title_pattern):
                move_window_to_monitor(hwnd, monitor_index, title)
                break
        else:
            exc = Exception(f"Janela '{title}' não foi movida para nenhum monitor.")
            logger.exception(exc)
            raise exc

if __name__ == "__main__":
    main()