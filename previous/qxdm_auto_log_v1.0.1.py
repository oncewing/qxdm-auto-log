import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import queue
import sys
import os
import time
import random
import subprocess
import json
from datetime import datetime
from collections import deque

__version__ = "1.0.1"
APP_NAME    = "QXDM AUTO LOG"

try:
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False


# ============================================================
# stdout → GUI 큐 라우터
# ============================================================
class GuiLogger:
    def __init__(self, text_queue, log_path):
        self.queue = text_queue
        self._terminal = sys.__stdout__
        self.log_file = open(log_path, "w", encoding="utf-8")

    def write(self, message):
        self.queue.put(message)
        self.log_file.write(message)
        self.log_file.flush()
        try:
            self._terminal.write(message)
        except Exception:
            pass

    def flush(self):
        self.log_file.flush()

    def close(self):
        try:
            self.log_file.close()
        except Exception:
            pass


# ============================================================
# 유틸
# ============================================================
def now_str():
    return datetime.now().strftime("%H:%M:%S.%f")[:-3]

def now_filestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


# ============================================================
# adb 헬퍼
# ============================================================
_DEVICE_ID = ""
_NO_WINDOW = subprocess.CREATE_NO_WINDOW  # cmd 창 팝업 방지

def _adb_base():
    return ["adb", "-s", _DEVICE_ID] if _DEVICE_ID else ["adb"]

def adb_devices():
    try:
        result = subprocess.run(["adb", "devices"], capture_output=True, text=True, timeout=5,
                                creationflags=_NO_WINDOW)
        devices = []
        for line in result.stdout.strip().splitlines()[1:]:
            line = line.strip()
            if line and "\t" in line:
                serial, state = line.split("\t", 1)
                if state.strip() == "device":
                    devices.append(serial.strip())
        return devices
    except Exception:
        return []

def adb_shell(cmd, timeout=10):
    try:
        result = subprocess.run(
            _adb_base() + ["shell", cmd],
            capture_output=True, text=True, timeout=timeout,
            creationflags=_NO_WINDOW,
        )
        return result.stdout.strip()
    except (subprocess.TimeoutExpired, FileNotFoundError) as e:
        print(f"  [adb 오류] {e}")
        return ""

def adb_at(cmd, timeout=10):
    """AT 명령 전송 및 응답 반환 (/data/shellat7 사용)"""
    return adb_shell(f"/data/shellat7 '{cmd}'", timeout=timeout)

def adb_check_device():
    return adb_shell("echo OK") == "OK"


# ============================================================
# 트리거 조건 실행 엔진
# ============================================================
DEFAULT_FUNC_CODE = """\
# 사용 가능: adb_shell(cmd), adb_at(cmd), time, stop_event
# 반환: (is_normal: bool, reason: str)
#   True  = 정상         → 테스트 계속
#   False = 비정상 감지  → QXDM 종료
# 주의: 명령 오류 메시지가 out에 섞이지 않도록 2>/dev/null 권장

out = adb_shell("cat /var/tmp/test_result 2>/dev/null | grep PASS")
if out:
    return True, "normal"    # 정상
return False, "error"        # 비정상
"""

DEFAULT_CONDITIONS = [
    {"type": "SHELL", "cmd": "cat /var/tmp/test_result", "expected": "error", "timeout": 15, "normal_when": "not_found"}    
]
DEFAULT_LOGIC = "AND"  # "AND" | "OR"


def _eval_cmd_condition(cond, stop_event, label, getter):
    """
    정상 조건 평가. 반환: (is_normal: bool, reason: str)
      normal_when="found"     : 기대값 출현 → 정상(True), 타임아웃 미출현 → 비정상(False)
      normal_when="not_found" : 기대값 미출현 유지 → 정상(True), 출현 즉시 → 비정상(False)
    """
    cmd         = cond.get("cmd", "")
    expected    = cond.get("expected", "")
    timeout     = int(cond.get("timeout", 10))
    normal_when = cond.get("normal_when", "found")  # "found" | "not_found"

    cond_desc = "출현 = 정상" if normal_when == "found" else "미출현 = 정상"
    print(f"  -> [{now_str()}] {label}: '{cmd}'  기대값='{expected}'  [{cond_desc}]  ({timeout}초)")

    start = time.time()
    while time.time() - start < timeout:
        if stop_event and stop_event.is_set():
            return False, "stopped"
        elapsed = int(time.time() - start)
        out     = getter(cmd)
        matched = expected in out

        if normal_when == "found" and matched:
            print(f"     [{now_str()}] ({elapsed}s) '{expected}' 출현 → 정상")
            return True, "normal"

        if normal_when == "not_found" and matched:
            print(f"     [{now_str()}] ({elapsed}s) '{expected}' 출현 → 비정상 → QXDM 종료")
            return False, f"unexpected: '{expected}' found"

        if elapsed % 3 == 0:
            print(f"     [{now_str()}] ({elapsed}s) '{expected}' {'있음' if matched else '없음'}...")
        time.sleep(1)

    # 타임아웃 처리
    if normal_when == "found":
        print(f"     [{now_str()}] {timeout}초 내 '{expected}' 미출현 → 비정상 → QXDM 종료")
        return False, f"timeout: '{expected}' not found"
    else:
        print(f"     [{now_str()}] {timeout}초간 '{expected}' 미출현 → 정상")
        return True, "normal"


def _eval_shell_condition(cond, stop_event, label):
    return _eval_cmd_condition(cond, stop_event, label, adb_shell)


def _eval_at_condition(cond, stop_event, label):
    return _eval_cmd_condition(cond, stop_event, label, adb_at)


def _eval_func_condition(cond, stop_event, label):
    """
    사용자 함수 평가. 함수는 (is_normal: bool, reason: str) 반환.
      True  = 정상         → 테스트 계속
      False = 비정상 감지  → QXDM 종료
    """
    code = cond.get("code", "return False, 'no_code'")
    func_src = "def _user_check(stop_event=None):\n"
    for line in code.splitlines():
        func_src += f"    {line}\n"
    print(f"  -> [{now_str()}] {label}: 사용자 함수 실행")
    try:
        ns = {"adb_shell": adb_shell, "adb_at": adb_at, "time": time}
        exec(func_src, ns)
        is_normal, reason = ns["_user_check"](stop_event=stop_event)
        is_normal = bool(is_normal)
        state = "정상" if is_normal else "비정상 → QXDM 종료"
        print(f"     [{now_str()}] 함수 반환: {state} ({reason})")
        return is_normal, str(reason)
    except Exception as e:
        import traceback as tb
        print(f"     [{now_str()}] 함수 실행 오류: {e}")
        tb.print_exc()
        return False, f"func_error: {e}"


def _eval_one(cond, stop_event, i):
    """조건 하나 평가. 반환: (is_normal: bool, reason: str)"""
    ctype = cond.get("type", "SHELL")
    label = f"조건 {i+1} [{ctype}]"
    if ctype == "SHELL":
        return _eval_shell_condition(cond, stop_event, label)
    elif ctype == "AT":
        return _eval_at_condition(cond, stop_event, label)
    elif ctype == "FUNCTION":
        return _eval_func_condition(cond, stop_event, label)
    return True, "skipped"


def run_conditions(conditions, stop_event=None, logic="AND"):
    """
    AND: 모든 조건이 정상(True)이어야 테스트 계속. 하나라도 False → QXDM 종료.
    OR : 조건 중 하나라도 정상(True)이면 테스트 계속. 전부 False → QXDM 종료.
    반환: (is_normal: bool, reason: str)
    """
    if logic == "AND":
        for i, cond in enumerate(conditions):
            if stop_event and stop_event.is_set():
                return False, "stopped"
            is_normal, reason = _eval_one(cond, stop_event, i)
            if not is_normal:
                return False, reason  # 하나라도 비정상 → 즉시 종료
        return True, "normal"

    else:  # OR
        failures = []
        for i, cond in enumerate(conditions):
            if stop_event and stop_event.is_set():
                return False, "stopped"
            is_normal, reason = _eval_one(cond, stop_event, i)
            if is_normal:
                return True, "normal"  # 하나라도 정상 → 계속
            failures.append(reason)
        return False, " | ".join(failures)


# ============================================================
# 테스트 단계
# ============================================================
DEFAULT_TEST_STEPS = [
    {"label": "DEFAULT",    "type": "AT",    "cmd": "AT",
     "base_sec": 5,  "rand_min": 1, "rand_max": 5}
]


def execute_test_step(step):
    """단계 명령 실행 (AT or SHELL)."""
    label    = step.get("label", "")
    cmd      = step.get("cmd", "")
    cmd_type = step.get("type", "AT")
    tag      = f"[{label}] " if label else ""
    if cmd_type == "AT":
        print(f"  -> [{now_str()}] {tag}AT 명령: {cmd}")
        adb_shell(f"/data/shellat7 '{cmd}' > /dev/null 2>&1")
    else:
        print(f"  -> [{now_str()}] {tag}Shell: {cmd}")
        adb_shell(cmd)

def force_crash_sysrq():
    """기존 방식: echo c > /proc/sysrq-trigger 로 커널 강제 crash"""
    print(f"  -> [{now_str()}] !!! 단말 강제 CRASH (sysrq-trigger) !!!")
    try:
        subprocess.run(["adb", "shell", "echo c > /proc/sysrq-trigger"],
                       capture_output=True, text=True, timeout=5,
                       creationflags=_NO_WINDOW)
    except subprocess.TimeoutExpired:
        print(f"     [{now_str()}] adb 응답 끊김 - crash 성공으로 추정")
    except Exception as e:
        print(f"     crash 명령 실행 오류: {e}")
    print(f"  -> [{now_str()}] crash 명령 전송 완료")


def _wait_device_reboot(stop_event=None, wait_timeout=120):
    """
    단말 재부팅 후 adb devices 에서 해당 device id 복귀 대기.
    반환: True = 복귀 확인, False = 타임아웃 / 중단
    """
    # 단말이 adb 에서 사라질 여유 시간
    print(f"     [{now_str()}] 재부팅 대기 중 (최대 {wait_timeout}초)...")
    time.sleep(8)

    start = time.time()
    while time.time() - start < wait_timeout:
        if stop_event and stop_event.is_set():
            print(f"     [{now_str()}] 중단 요청 - 단말 대기 종료")
            return False
        elapsed = int(time.time() - start)
        devices = adb_devices()
        if _DEVICE_ID:
            found = _DEVICE_ID in devices
        else:
            found = len(devices) > 0
        if found:
            target = _DEVICE_ID if _DEVICE_ID else devices[0]
            print(f"     [{now_str()}] ({elapsed}s) 단말 복귀 확인: {target}")
            time.sleep(3)   # 부팅 안정화 여유
            return True
        if elapsed % 10 == 0:
            print(f"     [{now_str()}] ({elapsed}s) 단말 대기 중... (adb: {devices})")
        time.sleep(2)

    print(f"     [{now_str()}] {wait_timeout}초 내 단말 미복귀")
    return False


def check_and_reboot_at(stop_event=None, wait_timeout=120):
    """
    AT!ERROROPT? 확인 후:
      - USB Dump 미포함 → AT!ERROROPT=0 + AT+CFUN=1,1 재부팅 → 단말 복귀 대기
      - USB Dump 포함   → sysrq-trigger 방식으로 처리 필요

    반환: (use_sysrq: bool, device_back: bool)
      use_sysrq=True  → USB Dump 포함, 호출자가 force_crash_sysrq() 실행
      use_sysrq=False, device_back=True  → AT 재부팅 완료, 단말 복귀 → 테스트 계속
      use_sysrq=False, device_back=False → AT 재부팅했으나 단말 미복귀 → 테스트 중단
    """
    print(f"  -> [{now_str()}] AT!ERROROPT? 확인 중...")
    erroropt = adb_at("AT!ERROROPT?", timeout=10)
    print(f"     응답: {erroropt!r}")

    if "USB Dump" in erroropt:
        print(f"     [USB Dump 포함] sysrq-trigger 방식으로 처리합니다.")
        return True, False

    # USB Dump 미포함 → ERROROPT 초기화 후 AT+CFUN=1,1 재부팅
    print(f"     [USB Dump 미포함] AT!ERROROPT=0 설정 중...")
    resp0 = adb_at("AT!ERROROPT=0", timeout=10)
    print(f"     AT!ERROROPT=0 응답: {resp0!r}")

    print(f"  -> [{now_str()}] AT+CFUN=1,1 재부팅 명령 전송...")
    try:
        adb_shell("/data/shellat7 'AT+CFUN=1,1' > /dev/null 2>&1", timeout=5)
    except Exception as e:
        print(f"     (AT+CFUN=1,1 응답 끊김 - 재부팅 진행 중으로 추정: {e})")
    print(f"  -> [{now_str()}] 재부팅 명령 전송 완료")

    device_back = _wait_device_reboot(stop_event=stop_event, wait_timeout=wait_timeout)
    return False, device_back

def sleep_with_progress(base_sec, rand_min, rand_max, label="", stop_event=None):
    rand_extra = random.randint(rand_min, rand_max)
    total = base_sec + rand_extra
    print(f"  -> {label} 대기: {base_sec}+random({rand_min}~{rand_max})={rand_extra} = 총 {total}초")
    for i in range(total):
        if stop_event and stop_event.is_set():
            print(f"     [중단] 대기 중 중단 요청됨")
            return total
        time.sleep(1)
        if (i + 1) % 5 == 0 or (i + 1) == total:
            print(f"     {i+1}/{total}s")
    return total


# ============================================================
# 저장 파일 탐지 / 롤링 저장
# ============================================================
LOG_EXTENSIONS = (".isf", ".hdf")

def wait_for_saved_file(save_dir, expected_path, before_files, max_wait=30):
    for _ in range(max_wait):
        time.sleep(1)
        if os.path.exists(expected_path):
            try:
                s1 = os.path.getsize(expected_path)
                time.sleep(0.5)
                if s1 == os.path.getsize(expected_path) and s1 > 0:
                    return expected_path
            except OSError:
                pass
        try:
            new = [f for f in set(os.listdir(save_dir)) - before_files
                   if f.lower().endswith(LOG_EXTENSIONS)]
            if new:
                cand = max([os.path.join(save_dir, f) for f in new], key=os.path.getmtime)
                s1 = os.path.getsize(cand)
                time.sleep(0.5)
                if s1 == os.path.getsize(cand) and s1 > 0:
                    return cand
        except Exception:
            pass
    return None


class RollingISFManager:
    def __init__(self, save_dir, base_prefix, keep_files):
        self.save_dir = save_dir
        self.base_prefix = base_prefix
        self.keep_files = keep_files
        self.saved_files = deque()
        self.rotation_index = 0

    def save_rotation(self, win):
        self.rotation_index += 1
        filename  = f"{self.base_prefix}_part{self.rotation_index:03d}_{now_filestamp()}.isf"
        filepath  = os.path.join(self.save_dir, filename)
        base_name = os.path.splitext(filename)[0]
        try:
            win.SetISFDirPath(self.save_dir)
            win.SetBaseISFFileName(base_name)
            time.sleep(0.5)
        except Exception as e:
            print(f"     (ISF 경로 설정 경고: {e})")
        item_count = -1
        try:
            item_count = win.GetItemCount()
        except Exception:
            pass
        before = set()
        try:
            before = set(os.listdir(self.save_dir))
        except Exception:
            pass
        print(f"     [{now_str()}] 롤링 저장 #{self.rotation_index}: {filename} (Item {item_count})")
        try:
            win.SaveItemStore(filepath)
        except Exception as e:
            print(f"     [경고] SaveItemStore 실패: {e}")
            return None
        actual = wait_for_saved_file(self.save_dir, filepath, before)
        if actual is None:
            print(f"     [경고] 30초 대기 후에도 저장 파일 없음")
            return None
        if actual != filepath:
            print(f"     -> 다른 이름으로 저장: {os.path.basename(actual)}")
        print(f"     -> 저장 완료 ({os.path.getsize(actual)/1024:.1f} KB): {os.path.basename(actual)}")
        self.saved_files.append(actual)
        while len(self.saved_files) > self.keep_files:
            old = self.saved_files.popleft()
            try:
                if os.path.exists(old):
                    os.remove(old)
                    print(f"     -> 오래된 파일 삭제: {os.path.basename(old)}")
            except Exception as e:
                print(f"     [경고] 파일 삭제 실패: {e}")
        return actual

    def list_kept_files(self):
        return list(self.saved_files)


def clear_qxdm_log(win):
    before = -1
    try:
        before = win.GetItemCount()
    except Exception:
        pass
    cleared = False
    for name in ("Item View", "ItemView", "Filtered View", "Default Item View"):
        try:
            win.ClearViewItems(name)
            cleared = True
            break
        except Exception:
            continue
    time.sleep(1)
    after = -1
    try:
        after = win.GetItemCount()
    except Exception:
        pass
    if cleared:
        print(f"     로그 클리어 완료 (Item 수: {before} -> {after})")
    else:
        print(f"     [경고] 로그 클리어 실패")
    return cleared


# ============================================================
# 메인 시나리오
# ============================================================
def run_test_scenario(
    save_dir, base_prefix, max_cycles,
    post_trigger_sec=5, rotate_every=0, keep_files=0,
    stop_event=None, phone_ready_event=None,
    poweroff=False, poweroff_delay=60,
    crash_on_trigger=False, device_id="",
    conditions=None, condition_logic="AND",
    test_steps=None,
):
    global _DEVICE_ID
    _DEVICE_ID = device_id
    if conditions is None:
        conditions = list(DEFAULT_CONDITIONS)
    if test_steps is None:
        test_steps = list(DEFAULT_TEST_STEPS)

    if not WIN32_AVAILABLE:
        print("[오류] win32com / pythoncom 패키지가 없습니다.")
        return

    pythoncom.CoInitialize()
    rolling = None
    try:
        if not adb_check_device():
            print("[오류] adb 기기를 찾을 수 없습니다.")
            return

        # ── 강제 Crash 옵션: QXDM 실행 전 AT!ERROROPT 확인 및 설정 ──────────
        if crash_on_trigger:
            print("=" * 60)
            print(f"[{now_str()}] [Crash 옵션] AT!ERROROPT 확인 및 초기화")
            print("=" * 60)
            use_sysrq, device_back = check_and_reboot_at(
                stop_event=stop_event, wait_timeout=120)
            if stop_event and stop_event.is_set():
                print("[중단] 사용자가 중단했습니다.")
                return
            if use_sysrq:
                # USB Dump 이미 포함 → 재부팅 불필요, 그대로 진행
                print(f"[{now_str()}] AT!ERROROPT 에 USB Dump 포함 확인 → 재부팅 없이 진행\n")
            elif device_back:
                print(f"[{now_str()}] 단말 복귀 완료 → QXDM 실행 진행\n")
            else:
                print(f"[{now_str()}] 단말 미복귀 → 테스트 중단")
                return
        # ───────────────────────────────────────────────────────────────────

        print("=" * 60)
        print(f"[{now_str()}] QXDM 시작 중...")
        print("=" * 60)
        app = win32com.client.Dispatch("QXDM.QXDM5AutoApplication")
        time.sleep(5)
        win = app.GetAutomationWindow2()
        win.SetVisible(True)
        try:
            win.SetAutoSaveISF(False)
        except Exception:
            pass
        print(f"QXDM 버전: {win.GetQXDMVersion()}")

        print("\nQXDM 창에서 Phone을 연결(Ctrl+O)한 뒤")
        print("GUI의 [✔ Phone 연결 완료] 버튼을 누르세요...")
        if phone_ready_event is not None:
            while not phone_ready_event.wait(timeout=0.5):
                if stop_event and stop_event.is_set():
                    print("[중단] 사용자가 중단했습니다.")
                    return
        else:
            input()

        if stop_event and stop_event.is_set():
            print("[중단] 사용자가 중단했습니다.")
            return

        if not win.GetIsPhoneConnected():
            print("Phone이 연결되지 않았습니다. 중단합니다.")
            win.QuitApplication()
            return
        print(f"[{now_str()}] Phone 연결 확인됨.\n")

        if rotate_every > 0:
            rolling = RollingISFManager(save_dir, base_prefix, keep_files)
            print(f"롤링 모드 활성: {rotate_every}회마다 저장, 최대 {keep_files}개 보관\n")

        triggered = False
        triggered_cycle = -1

        for cycle in range(1, max_cycles + 1):
            if stop_event and stop_event.is_set():
                print(f"\n[{now_str()}] 사용자 중단.")
                break

            print("\n" + "=" * 60)
            print(f"[{now_str()}] [Cycle {cycle}/{max_cycles}] 시작 (Item: {win.GetItemCount()})")
            print("=" * 60)
            try:
                win.AddAnnotationString(f"Cycle {cycle} START")
            except Exception:
                pass

            # 테스트 단계 실행
            step_aborted = False
            for step in test_steps:
                lbl = step.get("label", step.get("cmd", ""))
                execute_test_step(step)
                try:
                    win.AddAnnotationString(f"Cycle {cycle} {lbl}")
                except Exception:
                    pass
                sleep_with_progress(
                    step.get("base_sec", 5),
                    step.get("rand_min", 0),
                    step.get("rand_max", 0),
                    label=lbl, stop_event=stop_event,
                )
                if stop_event and stop_event.is_set():
                    step_aborted = True
                    break
            if step_aborted:
                break

            is_normal, reason = run_conditions(conditions, stop_event=stop_event,
                                               logic=condition_logic)
            if stop_event and stop_event.is_set():
                break

            if not is_normal:
                print(f"\n  >>> [{now_str()}] 트리거! (사유: {reason}) Cycle {cycle}")
                try:
                    win.AddAnnotationString(f"Cycle {cycle} TRIGGER ({reason})")
                except Exception:
                    pass
                triggered = True
                triggered_cycle = cycle
                if crash_on_trigger:
                    try:
                        win.AddAnnotationString(f"Cycle {cycle} CRASH")
                    except Exception:
                        pass
                    force_crash_sysrq()
                break
            else:
                print(f"     조건 정상 ({reason}) - 다음 사이클")

            if rolling and cycle % rotate_every == 0 and cycle != max_cycles:
                print(f"\n  -> [{now_str()}] 롤링 저장 (cycle {cycle})")
                try:
                    win.AddAnnotationString(f"Cycle {cycle} ROTATE")
                except Exception:
                    pass
                rolling.save_rotation(win)
                clear_qxdm_log(win)

        if not triggered and not (stop_event and stop_event.is_set()):
            print(f"\n[{now_str()}] {max_cycles}회 완료, 트리거 미발생.")

        if triggered and post_trigger_sec > 0:
            print(f"\n[{now_str()}] {post_trigger_sec}초 추가 수집...")
            for i in range(post_trigger_sec):
                if stop_event and stop_event.is_set():
                    break
                time.sleep(1)
                print(f"  {i+1}/{post_trigger_sec}s")

        print("\n" + "=" * 60)
        print(f"[{now_str()}] 최종 로그 저장")
        print("=" * 60)
        if triggered:
            final_name = f"{base_prefix}_TRIGGER_cycle{triggered_cycle:03d}_{now_filestamp()}.isf"
        else:
            final_name = f"{base_prefix}_FINAL_{now_filestamp()}.isf"
        final_path = os.path.join(save_dir, final_name)
        base_only  = os.path.splitext(final_name)[0]
        try:
            win.SetISFDirPath(save_dir)
            win.SetBaseISFFileName(base_only)
            time.sleep(1)
        except Exception as e:
            print(f"  (ISF 경로 설정 경고: {e})")

        print(f"총 수집 Item: {win.GetItemCount()}")
        if triggered:
            print(f"트리거 사이클: {triggered_cycle}")
        print(f"저장 경로: {final_path}")

        before = set()
        try:
            before = set(os.listdir(save_dir))
        except Exception:
            pass
        try:
            win.SaveItemStore(final_path)
        except Exception as e:
            print(f"  [경고] SaveItemStore 실패: {e}")

        actual = wait_for_saved_file(save_dir, final_path, before)
        if actual:
            if actual != final_path:
                print(f"  -> 다른 이름으로 저장: {os.path.basename(actual)}")
            print(f"저장 완료 ({os.path.getsize(actual)/1024:.1f} KB): {os.path.basename(actual)}")
        else:
            print("경고: 30초 대기 후에도 저장 파일 없음")

        if rolling:
            kept = rolling.list_kept_files()
            print(f"\n롤링 보관 파일 ({len(kept)}개):")
            for p in kept:
                kb = os.path.getsize(p) / 1024 if os.path.exists(p) else 0
                print(f"  - {os.path.basename(p)} ({kb:.1f} KB)")

        print(f"\n[{now_str()}] QXDM 종료 중...")
        win.QuitApplication()
        time.sleep(2)
        print(f"[{now_str()}] 완료.")

        if poweroff and not (stop_event and stop_event.is_set()):
            print(f"\n[!] {poweroff_delay}초 후 Windows 종료. 취소: shutdown /a")
            try:
                subprocess.run(["shutdown", "/s", "/t", str(poweroff_delay), "/f", "/d", "p:0:0"],
                               check=False, creationflags=_NO_WINDOW)
                print(f"[!] shutdown 명령 등록됨.")
            except Exception as e:
                print(f"[오류] shutdown 실패: {e}")

    except pythoncom.com_error as e:
        print(f"COM 오류: {e}")
    except Exception as e:
        import traceback
        print(f"오류: {e}")
        traceback.print_exc()
    finally:
        pythoncom.CoUninitialize()


# ============================================================
# 테스트 단계 편집 다이얼로그
# ============================================================
class TestStepEditDialog(tk.Toplevel):
    def __init__(self, parent, step=None):
        super().__init__(parent)
        self.title("테스트 단계 편집")
        self.resizable(False, False)
        self.grab_set()
        self.result = None
        self._step = step or {"label": "", "type": "AT", "cmd": "",
                               "base_sec": 5, "rand_min": 0, "rand_max": 0}
        self._build()
        self.transient(parent)
        self.wait_window()

    def _build(self):
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        f.columnconfigure(1, weight=1)

        def lbl_entry(row, text, val, width=35):
            ttk.Label(f, text=text).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=3)
            var = tk.StringVar(value=str(val))
            ttk.Entry(f, textvariable=var, width=width).grid(row=row, column=1, sticky="ew", pady=3)
            return var

        def lbl_spin(row, text, val, lo, hi):
            ttk.Label(f, text=text).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=3)
            var = tk.StringVar(value=str(val))
            ttk.Spinbox(f, from_=lo, to=hi, textvariable=var, width=8).grid(
                row=row, column=1, sticky="w", pady=3)
            return var

        s = self._step
        self.label_var = lbl_entry(0, "라벨 :", s.get("label", ""))

        ttk.Label(f, text="타입:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=3)
        self.type_var = tk.StringVar(value=s.get("type", "AT"))
        type_f = ttk.Frame(f)
        type_f.grid(row=1, column=1, sticky="w", pady=3)
        ttk.Radiobutton(type_f, text="AT", variable=self.type_var, value="AT").pack(side="left", padx=(0, 12))
        ttk.Radiobutton(type_f, text="SHELL", variable=self.type_var, value="SHELL").pack(side="left")

        self.cmd_var      = lbl_entry(2, "명령:", s.get("cmd", ""))
        self.base_var     = lbl_spin(3, "대기 기본(초):",      s.get("base_sec", 5),  0, 3600)
        self.rand_min_var = lbl_spin(4, "랜덤 추가 최소(초):", s.get("rand_min", 0),  0, 3600)
        self.rand_max_var = lbl_spin(5, "랜덤 추가 최대(초):", s.get("rand_max", 0),  0, 3600)

        ttk.Label(f, text="총 대기 = 기본 + random(최소 ~ 최대) 초",
                  foreground="gray").grid(row=6, column=0, columnspan=2, sticky="w", pady=(0, 6))

        btn = ttk.Frame(f)
        btn.grid(row=7, column=0, columnspan=2, sticky="e")
        ttk.Button(btn, text="취소", command=self.destroy, width=10).pack(side="right")
        ttk.Button(btn, text="확인", command=self._ok, width=10).pack(side="right", padx=(0, 4))

    def _ok(self):
        try:
            base = int(self.base_var.get())
            rmin = int(self.rand_min_var.get())
            rmax = int(self.rand_max_var.get())
        except ValueError:
            messagebox.showerror("입력 오류", "대기 시간은 정수여야 합니다.", parent=self)
            return
        if rmin > rmax:
            messagebox.showerror("입력 오류", "랜덤 최소 ≤ 최대 이어야 합니다.", parent=self)
            return
        self.result = {
            "label":    self.label_var.get().strip(),
            "type":     self.type_var.get(),
            "cmd":      self.cmd_var.get().strip(),
            "base_sec": base,
            "rand_min": rmin,
            "rand_max": rmax,
        }
        self.destroy()


# ============================================================
# 조건 편집 다이얼로그
# ============================================================
class ConditionEditDialog(tk.Toplevel):
    TYPES = ["SHELL", "AT", "FUNCTION"]

    def __init__(self, parent, cond=None):
        super().__init__(parent)
        self.title("조건 편집")
        self.grab_set()
        self.result = None
        self._cond = cond or {"type": "SHELL", "cmd": "", "expected": "", "timeout": 10}
        self._build()
        self.transient(parent)
        self.wait_window()

    def _build(self):
        # 타입 선택
        top = ttk.Frame(self, padding=(10, 10, 10, 4))
        top.pack(side="top", fill="x")
        ttk.Label(top, text="타입:").pack(side="left")
        self.type_var = tk.StringVar(value=self._cond.get("type", "SHELL"))
        cb = ttk.Combobox(top, textvariable=self.type_var, values=self.TYPES,
                          state="readonly", width=12)
        cb.pack(side="left", padx=6)
        cb.bind("<<ComboboxSelected>>", lambda _: self._switch_type())

        # 버튼 — body보다 먼저 side="bottom" 으로 pack해야 항상 하단에 고정됨
        btn = ttk.Frame(self, padding=(10, 6, 10, 10))
        btn.pack(side="bottom", fill="x")
        ttk.Button(btn, text="취소", command=self.destroy, width=10).pack(side="right")
        ttk.Button(btn, text="확인", command=self._ok, width=10).pack(side="right", padx=(0, 4))

        # 내용 영역 (버튼 다음에 pack — 남은 공간을 채움)
        self.body = ttk.Frame(self, padding=(10, 0, 10, 4))
        self.body.pack(side="top", fill="both", expand=True)

        self._switch_type()

    def _switch_type(self):
        for w in self.body.winfo_children():
            w.destroy()
        t = self.type_var.get()
        if t in ("SHELL", "AT"):
            self._build_cmd_form(t)
            self.minsize(580, 230)
            self.geometry("580x230")
        else:
            self._build_func_form()
            self.minsize(580, 480)
            self.geometry("580x480")

    def _lbl_entry(self, parent, label, value, row):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=3)
        var = tk.StringVar(value=str(value))
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", pady=3)
        parent.columnconfigure(1, weight=1)
        return var

    def _build_cmd_form(self, t):
        f = self.body
        lbl = "Shell 명령:" if t == "SHELL" else "AT 명령:"
        self.cmd_var      = self._lbl_entry(f, lbl,                    self._cond.get("cmd", ""),      0)
        self.expected_var = self._lbl_entry(f, "기대값 (포함 문자열):", self._cond.get("expected", ""), 1)

        # 정상 조건 (True/False)
        ttk.Label(f, text="정상 조건:").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=3)
        self.normal_when_var = tk.StringVar(value=self._cond.get("normal_when", "found"))
        tog = ttk.Frame(f)
        tog.grid(row=2, column=1, sticky="w", pady=3)
        ttk.Radiobutton(tog, text="출현 = 정상 (True)",
                        variable=self.normal_when_var, value="found").pack(side="left", padx=(0, 16))
        ttk.Radiobutton(tog, text="미출현 = 정상 (True)",
                        variable=self.normal_when_var, value="not_found").pack(side="left")

        ttk.Label(f, text="타임아웃(초):").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=3)
        self.timeout_var = tk.StringVar(value=str(self._cond.get("timeout", 10)))
        ttk.Spinbox(f, from_=1, to=3600, textvariable=self.timeout_var, width=8).grid(
            row=3, column=1, sticky="w", pady=3)

    def _build_func_form(self):
        f = self.body
        ttk.Label(f, text="반환: (True=정상 계속 / False=비정상→종료, reason: str)  |  사용: adb_shell(), adb_at()").pack(
            anchor="w", pady=(0, 4))
        self.func_text = scrolledtext.ScrolledText(
            f, font=("Consolas", 9), wrap="none",
            bg="#1e1e1e", fg="#d4d4d4", insertbackground="white",
        )
        self.func_text.pack(fill="both", expand=True)
        self.func_text.insert("1.0", self._cond.get("code", DEFAULT_FUNC_CODE))

    def _ok(self):
        t = self.type_var.get()
        if t in ("SHELL", "AT"):
            try:
                timeout = int(self.timeout_var.get())
            except ValueError:
                messagebox.showerror("입력 오류", "타임아웃은 정수여야 합니다.", parent=self)
                return
            self.result = {
                "type": t,
                "cmd": self.cmd_var.get().strip(),
                "expected": self.expected_var.get().strip(),
                "normal_when": self.normal_when_var.get(),
                "timeout": timeout,
            }
        else:
            self.result = {
                "type": "FUNCTION",
                "code": self.func_text.get("1.0", "end-1c"),
            }
        self.destroy()


# ============================================================
# 설정 저장/불러오기
# ============================================================
STEPS_CONFIG_FILE = "qxdm_auto_log_test_config.json"
COND_CONFIG_FILE  = "qxdm_auto_log_cond_config.json"
STEPS_TYPE_KEY    = "test_steps_config"
COND_TYPE_KEY     = "conditions_config"

def _base_dir():
    """EXE 옆 or 스크립트 옆 디렉터리"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def _steps_config_path():
    return os.path.join(_base_dir(), STEPS_CONFIG_FILE)

def _cond_config_path():
    return os.path.join(_base_dir(), COND_CONFIG_FILE)

def _load_json(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None

def _save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ============================================================
# GUI
# ============================================================
class ClatTestGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_NAME}  v{__version__}")
        self.root.minsize(720, 760)

        self.output_queue = queue.Queue()
        self.stop_event          = None
        self.phone_ready_event   = None
        self.test_thread         = None
        self.log_writer          = None

        # 실행 경로에 디폴트 config 파일이 있으면 자동 불러오기
        steps_cfg = _load_json(_steps_config_path())
        self.test_steps = (
            [dict(s) for s in steps_cfg["test_steps"]]
            if steps_cfg and steps_cfg.get("config_type") == STEPS_TYPE_KEY and "test_steps" in steps_cfg
            else [dict(s) for s in DEFAULT_TEST_STEPS]
        )
        cond_cfg = _load_json(_cond_config_path())
        self.conditions = (
            [dict(c) for c in cond_cfg["conditions"]]
            if cond_cfg and cond_cfg.get("config_type") == COND_TYPE_KEY and "conditions" in cond_cfg
            else [dict(c) for c in DEFAULT_CONDITIONS]
        )
        self._saved_logic = (
            cond_cfg.get("logic", DEFAULT_LOGIC)
            if cond_cfg and cond_cfg.get("config_type") == COND_TYPE_KEY
            else DEFAULT_LOGIC
        )

        self._build_ui()
        self._poll_output()

    # ── UI ─────────────────────────────────────────────────
    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        # 섹션 LabelFrame: 굵은 제목 + 진한 테두리
        style.configure(
            "Section.TLabelframe",
            bordercolor="#555555",
            lightcolor="#555555",
            darkcolor="#555555",
            borderwidth=2,
            relief="groove",
        )
        style.configure(
            "Section.TLabelframe.Label",
            font=("맑은 고딕", 10, "bold"),
            foreground="#1F497D",
        )

    def _build_ui(self):
        self._setup_styles()
        main = ttk.Frame(self.root, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(4, weight=1)

        self._build_params(main)
        self._build_test_steps_frame(main)
        self._build_conditions_frame(main)
        self._build_controls(main)
        self._build_log(main)

    def _build_params(self, parent):
        frame = ttk.LabelFrame(parent, text="테스트 파라미터", padding=8, style="Section.TLabelframe")
        frame.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        frame.columnconfigure(1, weight=1)

        def lbl(r, t):
            ttk.Label(frame, text=t).grid(row=r, column=0, sticky="w", pady=2, padx=(0, 8))

        def spin(r, var, lo, hi):
            ttk.Spinbox(frame, from_=lo, to=hi, textvariable=var, width=10).grid(
                row=r, column=1, sticky="w", padx=4)

        r = 0
        lbl(r, "Device:")
        self.device_var = tk.StringVar()
        self.device_combo = ttk.Combobox(frame, textvariable=self.device_var,
                                          width=32, state="readonly")
        self.device_combo.grid(row=r, column=1, sticky="w", padx=4)
        ttk.Button(frame, text="↺ 새로고침", command=self._refresh_devices).grid(row=r, column=2)
        self._refresh_devices()
        r += 1

        lbl(r, "저장 폴더:")
        self.save_dir_var = tk.StringVar(value=r"D:\QXDM_LOGS")
        ttk.Entry(frame, textvariable=self.save_dir_var).grid(row=r, column=1, sticky="ew", padx=4)
        ttk.Button(frame, text="찾아보기", command=self._browse_dir).grid(row=r, column=2)
        r += 1

        # 사이클 수 + 트리거 후 수집 + 로그 저장 주기 + 로그 보관 개수 — 한 라인
        inline1 = ttk.Frame(frame)
        inline1.grid(row=r, column=0, columnspan=3, sticky="w", pady=2)
        ttk.Label(inline1, text="사이클 수:").pack(side="left")
        self.cycles_var = tk.IntVar(value=100)
        ttk.Spinbox(inline1, from_=1, to=9999, textvariable=self.cycles_var, width=5).pack(side="left", padx=(4, 16))
        ttk.Label(inline1, text="트리거 후 수집(초):").pack(side="left")
        self.post_sec_var = tk.IntVar(value=5)
        ttk.Spinbox(inline1, from_=0, to=3600, textvariable=self.post_sec_var, width=3).pack(side="left", padx=(4, 16))
        ttk.Label(inline1, text="로그 저장 주기:").pack(side="left")
        self.rotate_every_var = tk.IntVar(value=1)
        ttk.Spinbox(inline1, from_=0, to=9999, textvariable=self.rotate_every_var, width=3).pack(side="left", padx=(4, 16))
        ttk.Label(inline1, text="로그 보관 개수:").pack(side="left")
        self.keep_files_var = tk.IntVar(value=5)
        ttk.Spinbox(inline1, from_=1, to=100, textvariable=self.keep_files_var, width=3).pack(side="left", padx=4)
        r += 1

        # 강제 Crash + 테스트 후 Windows 종료 + 종료 대기 — 한 라인
        inline3 = ttk.Frame(frame)
        inline3.grid(row=r, column=0, columnspan=3, sticky="w", pady=2)
        self.crash_on_trigger_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(inline3, text="트리거 발생 시 강제 Crash",
                        variable=self.crash_on_trigger_var).pack(side="left", padx=(0, 20))
        self.poweroff_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(inline3, text="테스트 후 Windows 종료",
                        variable=self.poweroff_var,
                        command=self._toggle_poweroff).pack(side="left", padx=(0, 8))
        ttk.Label(inline3, text="종료 대기(초):").pack(side="left")
        self.poweroff_delay_var = tk.IntVar(value=60)
        self.poweroff_delay_spin = ttk.Spinbox(
            inline3, from_=0, to=3600, textvariable=self.poweroff_delay_var,
            width=4, state="disabled")
        self.poweroff_delay_spin.pack(side="left", padx=4)

    def _build_test_steps_frame(self, parent):
        self.step_outer = ttk.LabelFrame(parent, text="테스트 단계  (순서대로 실행)", padding=8, style="Section.TLabelframe")
        self.step_outer.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        self.step_outer.columnconfigure(0, weight=1)
        self._refresh_steps_ui()

    def _refresh_steps_ui(self):
        for w in self.step_outer.winfo_children():
            w.destroy()
        for i, step in enumerate(self.test_steps):
            row_f = ttk.Frame(self.step_outer)
            row_f.grid(row=i, column=0, sticky="ew", pady=1)
            row_f.columnconfigure(2, weight=1)
            ttk.Label(row_f, text=f"#{i+1}", width=3).grid(row=0, column=0)
            ttk.Label(row_f, text=f"[{step['type']}]", width=8,
                      foreground="#569cd6").grid(row=0, column=1, sticky="w")
            ttk.Label(row_f, text=self._step_summary(step),
                      anchor="w").grid(row=0, column=2, sticky="ew", padx=6)
            ttk.Button(row_f, text="편집", width=6,
                       command=lambda idx=i: self._edit_step(idx)).grid(row=0, column=3, padx=2)
            ttk.Button(row_f, text="삭제", width=6,
                       command=lambda idx=i: self._delete_step(idx)).grid(row=0, column=4)
        add_row = ttk.Frame(self.step_outer)
        add_row.grid(row=len(self.test_steps), column=0, sticky="ew", pady=(6, 0))
        ttk.Button(add_row, text="↺ 초기화",    command=self._reset_steps).pack(side="left", padx=(0, 4))
        ttk.Button(add_row, text="💾 저장",      command=self._save_steps_dialog).pack(side="left", padx=(0, 4))
        ttk.Button(add_row, text="📂 불러오기",  command=self._load_steps_dialog).pack(side="left")
        ttk.Button(add_row, text="+ 단계 추가",  command=self._add_step).pack(side="right")

    def _step_summary(self, step):
        lbl  = f"[{step['label']}]  " if step.get("label") else ""
        cmd  = step.get("cmd", "")[:40]
        base = step.get("base_sec", 0)
        rmin = step.get("rand_min", 0)
        rmax = step.get("rand_max", 0)
        return f"{lbl}{cmd}   대기: {base}+rand({rmin}~{rmax})초"

    def _add_step(self):
        dlg = TestStepEditDialog(self.root)
        if dlg.result:
            self.test_steps.append(dlg.result)
            self._refresh_steps_ui()

    def _edit_step(self, idx):
        dlg = TestStepEditDialog(self.root, self.test_steps[idx])
        if dlg.result:
            self.test_steps[idx] = dlg.result
            self._refresh_steps_ui()

    def _delete_step(self, idx):
        if len(self.test_steps) <= 1:
            messagebox.showwarning("경고", "단계는 최소 1개 이상이어야 합니다.", parent=self.root)
            return
        self.test_steps.pop(idx)
        self._refresh_steps_ui()

    def _reset_steps(self):
        if messagebox.askyesno("초기화", "테스트 단계를 기본값으로 초기화하시겠습니까?", parent=self.root):
            self.test_steps = [dict(s) for s in DEFAULT_TEST_STEPS]
            self._refresh_steps_ui()

    def _build_conditions_frame(self, parent):
        self.cond_outer = ttk.LabelFrame(
            parent, text="정상 조건  (마지막 단계 후 체크 — False 시 QXDM 종료)", padding=8,
            style="Section.TLabelframe")
        self.cond_outer.grid(row=2, column=0, sticky="ew", pady=(0, 6))
        self.cond_outer.columnconfigure(0, weight=1)

        # AND / OR 선택
        logic_row = ttk.Frame(self.cond_outer)
        logic_row.grid(row=0, column=0, sticky="w", pady=(0, 4))
        ttk.Label(logic_row, text="조건 로직:").pack(side="left", padx=(0, 8))
        self.condition_logic_var = tk.StringVar(value=getattr(self, "_saved_logic", DEFAULT_LOGIC))
        ttk.Radiobutton(logic_row, text="AND  (모두 정상이어야 계속)",
                        variable=self.condition_logic_var, value="AND").pack(side="left", padx=(0, 12))
        ttk.Radiobutton(logic_row, text="OR  (하나만 정상이어도 계속)",
                        variable=self.condition_logic_var, value="OR").pack(side="left")

        ttk.Separator(self.cond_outer, orient="horizontal").grid(
            row=1, column=0, sticky="ew", pady=(0, 4))

        self._cond_list_row = 2  # 조건 목록 시작 row
        self._refresh_conditions_ui()

    def _refresh_conditions_ui(self):
        # AND/OR 행(row 0,1)은 유지, 조건 목록(row 2+)만 갱신
        base = self._cond_list_row
        for w in self.cond_outer.grid_slaves():
            if int(w.grid_info().get("row", 0)) >= base:
                w.destroy()

        for i, cond in enumerate(self.conditions):
            row_f = ttk.Frame(self.cond_outer)
            row_f.grid(row=base + i, column=0, sticky="ew", pady=1)
            row_f.columnconfigure(2, weight=1)

            ttk.Label(row_f, text=f"{i+1}.", width=3).grid(row=0, column=0)
            ttk.Label(row_f, text=f"[{cond['type']}]", width=10,
                      foreground="#569cd6").grid(row=0, column=1, sticky="w")
            ttk.Label(row_f, text=self._cond_summary(cond),
                      anchor="w").grid(row=0, column=2, sticky="ew", padx=6)
            ttk.Button(row_f, text="편집", width=6,
                       command=lambda idx=i: self._edit_condition(idx)).grid(row=0, column=3, padx=2)
            ttk.Button(row_f, text="삭제", width=6,
                       command=lambda idx=i: self._delete_condition(idx)).grid(row=0, column=4)

        add_row = ttk.Frame(self.cond_outer)
        add_row.grid(row=base + len(self.conditions), column=0, sticky="ew", pady=(6, 0))
        ttk.Button(add_row, text="↺ 초기화",    command=self._reset_conditions).pack(side="left", padx=(0, 4))
        ttk.Button(add_row, text="💾 저장",      command=self._save_cond_dialog).pack(side="left", padx=(0, 4))
        ttk.Button(add_row, text="📂 불러오기",  command=self._load_cond_dialog).pack(side="left")
        ttk.Button(add_row, text="+ 조건 추가",  command=self._add_condition).pack(side="right")

    def _cond_summary(self, cond):
        t = cond.get("type", "SHELL")
        if t in ("SHELL", "AT"):
            cmd         = cond.get("cmd", "")[:40]
            exp         = cond.get("expected", "")[:20]
            normal_when = cond.get("normal_when", "found")
            flag        = "[출현=정상]" if normal_when == "found" else "[미출현=정상]"
            return f"{cmd}  →  '{exp}'  {flag}  ({cond.get('timeout', '?')}s)"
        first = (cond.get("code", "").strip().splitlines() or [""])[0][:60]
        return f"func: {first}"

    def _add_condition(self):
        dlg = ConditionEditDialog(self.root)
        if dlg.result:
            self.conditions.append(dlg.result)
            self._refresh_conditions_ui()

    def _edit_condition(self, idx):
        dlg = ConditionEditDialog(self.root, self.conditions[idx])
        if dlg.result:
            self.conditions[idx] = dlg.result
            self._refresh_conditions_ui()

    def _delete_condition(self, idx):
        if len(self.conditions) <= 1:
            messagebox.showwarning("경고", "조건은 최소 1개 이상이어야 합니다.", parent=self.root)
            return
        self.conditions.pop(idx)
        self._refresh_conditions_ui()

    def _reset_conditions(self):
        if messagebox.askyesno("초기화", "정상 조건을 기본값으로 초기화하시겠습니까?", parent=self.root):
            self.conditions = [dict(c) for c in DEFAULT_CONDITIONS]
            self.condition_logic_var.set(DEFAULT_LOGIC)
            self._refresh_conditions_ui()

    def _build_controls(self, parent):
        frame = ttk.Frame(parent)
        frame.grid(row=3, column=0, sticky="ew", pady=6)

        self.start_btn = ttk.Button(frame, text="▶  START", command=self._start_test, width=14)
        self.start_btn.pack(side="left", padx=(0, 6))

        self.stop_btn = ttk.Button(frame, text="■  STOP", command=self._stop_test,
                                   state="disabled", width=14)
        self.stop_btn.pack(side="left", padx=(0, 6))

        self.phone_btn = ttk.Button(frame, text="✔  Phone 연결 완료",
                                    command=self._phone_ready, state="disabled", width=20)
        self.phone_btn.pack(side="left", padx=(12, 0))

        self.status_var = tk.StringVar(value="대기 중")
        ttk.Label(frame, textvariable=self.status_var, foreground="gray",
                  width=12, anchor="e").pack(side="right")

    def _build_log(self, parent):
        frame = ttk.LabelFrame(parent, text="로그 출력", padding=4, style="Section.TLabelframe")
        frame.grid(row=4, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(frame, wrap="word", state="disabled",
                                font=("Consolas", 9), bg="#F0F0F0", fg="#1a1a1a",
                                insertbackground="black")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        sb = ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=sb.set)
        ttk.Button(frame, text="로그 지우기", command=self._clear_log).grid(
            row=1, column=0, sticky="e", pady=(4, 0))

        self.log_text.tag_configure("error",   foreground="#CC0000")
        self.log_text.tag_configure("trigger", foreground="#E65C00")
        self.log_text.tag_configure("ok",      foreground="#007A5E")
        self.log_text.tag_configure("header",  foreground="#1F497D")
        self.log_text.tag_configure("warn",    foreground="#8B6914")

    # ── 동작 ────────────────────────────────────────────────
    def _refresh_devices(self):
        devices = adb_devices()
        self.device_combo["values"] = devices
        if devices and self.device_var.get() not in devices:
            self.device_var.set(devices[0])
        elif not devices:
            self.device_var.set("")

    def _toggle_poweroff(self):
        self.poweroff_delay_spin.configure(
            state="normal" if self.poweroff_var.get() else "disabled")

    def _browse_dir(self):
        d = filedialog.askdirectory(initialdir=self.save_dir_var.get())
        if d:
            self.save_dir_var.set(d)

    def _append_log(self, text: str):
        self.log_text.configure(state="normal")
        if any(k in text for k in ("[오류]", "오류:", "ERROR", "com_error")):
            tag = "error"
        elif any(k in text for k in ("트리거", "TRIGGER", "이상 상태")):
            tag = "trigger"
        elif any(k in text for k in ("[경고]", "경고:")):
            tag = "warn"
        elif any(k in text for k in ("완료", "정상", "확인됨")):
            tag = "ok"
        elif text.startswith("="):
            tag = "header"
        else:
            tag = ""
        self.log_text.insert("end", text, tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _poll_output(self):
        try:
            while True:
                self._append_log(self.output_queue.get_nowait())
        except queue.Empty:
            pass
        self.root.after(50, self._poll_output)

    def _validate(self):
        try:
            if self.cycles_var.get() < 1:
                raise ValueError
        except (ValueError, tk.TclError):
            messagebox.showerror("입력 오류", "사이클 수는 1 이상이어야 합니다.")
            return False
        if self.rotate_every_var.get() > 0 and self.keep_files_var.get() < 1:
            messagebox.showerror("입력 오류", "롤링 모드에서 보관 개수는 1 이상이어야 합니다.")
            return False
        if not self.conditions:
            messagebox.showerror("입력 오류", "트리거 조건이 최소 1개 이상 필요합니다.")
            return False
        return True

    def _start_test(self):
        if not self._validate():
            return
        save_dir = self.save_dir_var.get()
        if not os.path.exists(save_dir):
            try:
                os.makedirs(save_dir)
            except Exception as e:
                messagebox.showerror("오류", f"저장 폴더 생성 실패:\n{e}")
                return

        self.stop_event        = threading.Event()
        self.phone_ready_event = threading.Event()

        start_ts    = now_filestamp()
        base_prefix = f"test_{start_ts}"
        log_path    = os.path.join(save_dir, f"{base_prefix}.log")

        self.log_writer = GuiLogger(self.output_queue, log_path)
        sys.stdout = self.log_writer

        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.phone_btn.configure(state="normal")
        self.status_var.set("실행 중...")

        print(f"테스트 시작: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Device: {self.device_var.get() or '(기본)'}")
        print(f"사이클 수: {self.cycles_var.get()}")
        print(f"트리거 후 수집: {self.post_sec_var.get()}초")
        print(f"롤링 저장 주기: {self.rotate_every_var.get() or '없음'}")
        if self.rotate_every_var.get() > 0:
            print(f"롤링 보관 개수: 최대 {self.keep_files_var.get()}개")
        print(f"트리거 시 강제 Crash: {'예 (시작 시 AT!ERROROPT 확인 + 재부팅)' if self.crash_on_trigger_var.get() else '아니오'}")
        print(f"테스트 후 Windows 종료: {'예' if self.poweroff_var.get() else '아니오'}")
        print(f"저장 폴더: {save_dir}")
        print(f"로그 파일: {log_path}")
        print(f"정상 조건 ({len(self.conditions)}개, 로직: {self.condition_logic_var.get()}):")
        for i, c in enumerate(self.conditions):
            print(f"  {i+1}. [{c['type']}] {self._cond_summary(c)}")
        print()

        params = dict(
            save_dir=save_dir, base_prefix=base_prefix,
            max_cycles=self.cycles_var.get(),
            post_trigger_sec=self.post_sec_var.get(),
            rotate_every=self.rotate_every_var.get(),
            keep_files=self.keep_files_var.get(),
            stop_event=self.stop_event,
            phone_ready_event=self.phone_ready_event,
            poweroff=self.poweroff_var.get(),
            poweroff_delay=self.poweroff_delay_var.get(),
            crash_on_trigger=self.crash_on_trigger_var.get(),
            device_id=self.device_var.get(),
            conditions=[dict(c) for c in self.conditions],
            condition_logic=self.condition_logic_var.get(),
            test_steps=[dict(s) for s in self.test_steps],
        )
        self.test_thread = threading.Thread(target=self._run_thread, kwargs=params, daemon=True)
        self.test_thread.start()

    def _run_thread(self, **kwargs):
        try:
            run_test_scenario(**kwargs)
        except Exception as e:
            import traceback
            print(f"\n[GUI] 예외: {e}")
            traceback.print_exc()
        finally:
            self.root.after(0, self._on_test_done)

    def _on_test_done(self):
        sys.stdout = sys.__stdout__
        if self.log_writer:
            self.log_writer.close()
            self.log_writer = None
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.phone_btn.configure(state="disabled")
        self.status_var.set("중단됨" if self.stop_event and self.stop_event.is_set() else "완료")

    def _stop_test(self):
        if self.stop_event:
            self.stop_event.set()
        if self.phone_ready_event:
            self.phone_ready_event.set()
        self.status_var.set("중단 중...")
        self.stop_btn.configure(state="disabled")

    def _phone_ready(self):
        if self.phone_ready_event:
            self.phone_ready_event.set()
        self.phone_btn.configure(state="disabled")
        self.output_queue.put("\n[GUI] Phone 연결 완료 확인 → 테스트 계속 진행\n")

    # ── 테스트 단계 저장/불러오기 ───────────────────────────
    def _save_steps_dialog(self):
        path = filedialog.asksaveasfilename(
            title="테스트 단계 저장",
            initialdir=_base_dir(),
            initialfile=STEPS_CONFIG_FILE,
            defaultextension=".json",
            filetypes=[("JSON 파일", "*.json"), ("모든 파일", "*.*")],
            parent=self.root,
        )
        if not path:
            return
        try:
            _save_json(path, {
                "config_type": STEPS_TYPE_KEY,
                "test_steps":  self.test_steps,
            })
            messagebox.showinfo("저장 완료", f"테스트 단계가 저장되었습니다.\n{path}", parent=self.root)
        except Exception as e:
            messagebox.showerror("저장 실패", str(e), parent=self.root)

    def _load_steps_dialog(self):
        path = filedialog.askopenfilename(
            title="테스트 단계 불러오기",
            initialdir=_base_dir(),
            initialfile=STEPS_CONFIG_FILE,
            filetypes=[("JSON 파일", "*.json"), ("모든 파일", "*.*")],
            parent=self.root,
        )
        if not path:
            return
        try:
            data = _load_json(path)
            if data is None:
                raise ValueError("파일을 읽을 수 없습니다.")
            if data.get("config_type") != STEPS_TYPE_KEY:
                messagebox.showwarning(
                    "파일 형식 불일치",
                    f"선택한 파일은 테스트 단계 설정 파일이 아닙니다.\n"
                    f"(config_type: {data.get('config_type', '없음')})\n\n"
                    f"테스트 단계 파일({STEPS_CONFIG_FILE})을 선택하세요.",
                    parent=self.root,
                )
                return
            self.test_steps = [dict(s) for s in data["test_steps"]]
            self._refresh_steps_ui()
            messagebox.showinfo("불러오기 완료", f"테스트 단계를 불러왔습니다.\n{path}", parent=self.root)
        except Exception as e:
            messagebox.showerror("불러오기 실패", str(e), parent=self.root)

    # ── 정상 조건 저장/불러오기 ─────────────────────────────
    def _save_cond_dialog(self):
        path = filedialog.asksaveasfilename(
            title="정상 조건 저장",
            initialdir=_base_dir(),
            initialfile=COND_CONFIG_FILE,
            defaultextension=".json",
            filetypes=[("JSON 파일", "*.json"), ("모든 파일", "*.*")],
            parent=self.root,
        )
        if not path:
            return
        try:
            _save_json(path, {
                "config_type": COND_TYPE_KEY,
                "logic":       self.condition_logic_var.get(),
                "conditions":  self.conditions,
            })
            messagebox.showinfo("저장 완료", f"정상 조건이 저장되었습니다.\n{path}", parent=self.root)
        except Exception as e:
            messagebox.showerror("저장 실패", str(e), parent=self.root)

    def _load_cond_dialog(self):
        path = filedialog.askopenfilename(
            title="정상 조건 불러오기",
            initialdir=_base_dir(),
            initialfile=COND_CONFIG_FILE,
            filetypes=[("JSON 파일", "*.json"), ("모든 파일", "*.*")],
            parent=self.root,
        )
        if not path:
            return
        try:
            data = _load_json(path)
            if data is None:
                raise ValueError("파일을 읽을 수 없습니다.")
            if data.get("config_type") != COND_TYPE_KEY:
                messagebox.showwarning(
                    "파일 형식 불일치",
                    f"선택한 파일은 정상 조건 설정 파일이 아닙니다.\n"
                    f"(config_type: {data.get('config_type', '없음')})\n\n"
                    f"정상 조건 파일({COND_CONFIG_FILE})을 선택하세요.",
                    parent=self.root,
                )
                return
            self.conditions = [dict(c) for c in data["conditions"]]
            self.condition_logic_var.set(data.get("logic", DEFAULT_LOGIC))
            self._refresh_conditions_ui()
            messagebox.showinfo("불러오기 완료", f"정상 조건을 불러왔습니다.\n{path}", parent=self.root)
        except Exception as e:
            messagebox.showerror("불러오기 실패", str(e), parent=self.root)


# ============================================================
# 진입점
# ============================================================
def main():
    root = tk.Tk()
    ClatTestGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
