
import csv
import ctypes
import io
import json
import random
import subprocess
import sys
import threading
import time
from collections import deque
from copy import deepcopy
from dataclasses import asdict, dataclass, field
from pathlib import Path

import winsound
import win32api
import win32con

try:
    from PySide6.QtCore import Qt, QTimer, Signal, QObject
    from PySide6.QtWidgets import (
        QApplication,
        QCheckBox,
        QComboBox,
        QDoubleSpinBox,
        QFileDialog,
        QFormLayout,
        QFrame,
        QGridLayout,
        QGroupBox,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMainWindow,
        QPushButton,
        QSpinBox,
        QTextEdit,
        QVBoxLayout,
        QWidget,
    )
except Exception:
    print("缺少 PySide6，請先安裝：pip install PySide6")
    raise SystemExit(1)


DEFAULT_TRIGGER_VK = win32con.VK_XBUTTON2
HOTKEY_SCAN_VKS = list(range(1, 255))


@dataclass
class DirectionConfig:
    enabled: bool = True
    mode: str = "squeeze"  # squeeze/drop
    random_hold_min_ms: float = 4500.0
    random_hold_max_ms: float = 7000.0
    replay_jitter_min_ms: float = 0.0
    replay_jitter_max_ms: float = 0.0
    random_bw_min_kbps: float = 120.0
    random_bw_max_kbps: float = 1200.0
    random_bw_resample_ms: float = 300.0
    replay_loss_rate_pct: float = 0.0
    drop_rate_pct: float = 100.0


@dataclass
class AppConfig:
    target_exe: str = ""
    windivert_dll: str = "WinDivert.dll"
    trigger_vk: int = DEFAULT_TRIGGER_VK
    trigger_mode: str = "toggle"  # toggle/hold
    enable_left_click_restore: bool = False
    poll_interval_ms: int = 10
    max_filter_ports: int = 200
    max_packet_size: int = 65535
    beep_frequency: int = 880
    beep_duration_ms: int = 180
    restore_beep_frequency: int = 660
    restore_beep_duration_ms: int = 220
    outbound: DirectionConfig = field(default_factory=DirectionConfig)
    inbound: DirectionConfig = field(default_factory=lambda: DirectionConfig(enabled=False))

    @classmethod
    def from_dict(cls, data):
        cfg = cls()
        if not isinstance(data, dict):
            return cfg

        def gi(k, d):
            try:
                return int(data.get(k, d))
            except Exception:
                return d

        def gf(v, d):
            try:
                return float(v)
            except Exception:
                return d

        cfg.target_exe = str(data.get("target_exe", cfg.target_exe))
        cfg.windivert_dll = str(data.get("windivert_dll", cfg.windivert_dll))
        cfg.trigger_vk = gi("trigger_vk", cfg.trigger_vk)
        cfg.trigger_mode = str(data.get("trigger_mode", cfg.trigger_mode)).lower()
        if cfg.trigger_mode not in {"toggle", "hold"}:
            cfg.trigger_mode = "toggle"
        cfg.enable_left_click_restore = bool(data.get("enable_left_click_restore", cfg.enable_left_click_restore))
        cfg.poll_interval_ms = max(1, gi("poll_interval_ms", cfg.poll_interval_ms))
        cfg.max_filter_ports = max(1, gi("max_filter_ports", cfg.max_filter_ports))
        cfg.max_packet_size = max(512, gi("max_packet_size", cfg.max_packet_size))
        cfg.beep_frequency = max(37, gi("beep_frequency", cfg.beep_frequency))
        cfg.beep_duration_ms = max(10, gi("beep_duration_ms", cfg.beep_duration_ms))
        cfg.restore_beep_frequency = max(37, gi("restore_beep_frequency", cfg.restore_beep_frequency))
        cfg.restore_beep_duration_ms = max(10, gi("restore_beep_duration_ms", cfg.restore_beep_duration_ms))

        def load_dir(raw, default_dir):
            d = deepcopy(default_dir)
            if not isinstance(raw, dict):
                return d
            d.enabled = bool(raw.get("enabled", d.enabled))
            d.mode = str(raw.get("mode", d.mode)).lower()
            if d.mode not in {"squeeze", "drop"}:
                d.mode = "squeeze"
            d.random_hold_min_ms = gf(raw.get("random_hold_min_ms", d.random_hold_min_ms), d.random_hold_min_ms)
            d.random_hold_max_ms = gf(raw.get("random_hold_max_ms", d.random_hold_max_ms), d.random_hold_max_ms)
            d.replay_jitter_min_ms = gf(raw.get("replay_jitter_min_ms", d.replay_jitter_min_ms), d.replay_jitter_min_ms)
            d.replay_jitter_max_ms = gf(raw.get("replay_jitter_max_ms", d.replay_jitter_max_ms), d.replay_jitter_max_ms)
            d.random_bw_min_kbps = gf(raw.get("random_bw_min_kbps", d.random_bw_min_kbps), d.random_bw_min_kbps)
            d.random_bw_max_kbps = gf(raw.get("random_bw_max_kbps", d.random_bw_max_kbps), d.random_bw_max_kbps)
            d.random_bw_resample_ms = gf(raw.get("random_bw_resample_ms", d.random_bw_resample_ms), d.random_bw_resample_ms)
            d.replay_loss_rate_pct = gf(raw.get("replay_loss_rate_pct", d.replay_loss_rate_pct), d.replay_loss_rate_pct)
            d.drop_rate_pct = gf(raw.get("drop_rate_pct", d.drop_rate_pct), d.drop_rate_pct)
            return d

        cfg.outbound = load_dir(data.get("outbound"), cfg.outbound)
        cfg.inbound = load_dir(data.get("inbound"), cfg.inbound)
        return cfg


def clamp_pct(v):
    try:
        x = float(v)
    except Exception:
        x = 0.0
    return max(0.0, min(100.0, x))


def norm_range(a, b):
    lo = max(0.0, float(a))
    hi = max(0.0, float(b))
    if hi < lo:
        lo, hi = hi, lo
    return lo, hi


def kbps_to_bps(kbps):
    if kbps <= 0:
        return 0.0
    return (kbps * 1000.0) / 8.0


def run_cmd(args):
    r = subprocess.run(args, capture_output=True, text=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    return r.returncode, r.stdout, r.stderr


def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def resolve_windivert_dll_path():
    # Prefer dll next to running script/executable.
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).resolve().parent
    else:
        base = Path(__file__).resolve().parent
    primary = base / "WinDivert.dll"
    if primary.exists():
        return primary

    # Fallback for some bundled runtimes.
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        fallback = Path(meipass) / "WinDivert.dll"
        if fallback.exists():
            return fallback

    return primary


def safe_beep(freq, duration_ms, fallback):
    try:
        winsound.Beep(max(37, int(freq)), max(10, int(duration_ms)))
    except RuntimeError:
        winsound.MessageBeep(fallback)
    except Exception:
        pass


def parse_port(local_endpoint):
    endpoint = local_endpoint.strip()
    if endpoint in {"*", "*:*"}:
        return None
    if endpoint.startswith("[") and "]:" in endpoint:
        p = endpoint.rsplit("]:", 1)[-1]
    else:
        if ":" not in endpoint:
            return None
        p = endpoint.rsplit(":", 1)[-1]
    if not p.isdigit():
        return None
    v = int(p)
    if 0 <= v <= 65535:
        return v
    return None


def get_target_pids(exe_path):
    image_name = Path(exe_path).name
    code, out, err = run_cmd(["tasklist", "/FI", f"IMAGENAME eq {image_name}", "/FO", "CSV", "/NH"])
    if code != 0:
        return set(), err.strip()
    pids = set()
    for row in csv.reader(io.StringIO(out)):
        if len(row) < 2:
            continue
        if row[0].strip().strip('"').lower() != image_name.lower():
            continue
        pid_text = row[1].strip().strip('"').replace(",", "")
        if pid_text.isdigit():
            pids.add(int(pid_text))
    return pids, ""


def collect_ports_for_pids(target_pids):
    tcp, udp = set(), set()
    for proto in ("tcp", "udp"):
        code, out, _ = run_cmd(["netstat", "-ano", "-p", proto])
        if code != 0:
            continue
        for raw in out.splitlines():
            line = raw.strip()
            if not line:
                continue
            parts = line.split()
            if not parts:
                continue
            h = parts[0].upper()
            if h == "TCP" and proto == "tcp" and len(parts) >= 5:
                local_ep, pid_text = parts[1], parts[-1]
            elif h == "UDP" and proto == "udp" and len(parts) >= 4:
                local_ep, pid_text = parts[1], parts[-1]
            else:
                continue
            if not pid_text.isdigit() or int(pid_text) not in target_pids:
                continue
            p = parse_port(local_ep)
            if p is None:
                continue
            (tcp if proto == "tcp" else udp).add(p)
    return tcp, udp


def build_filter(direction_kw, tcp_ports, udp_ports, max_ports):
    terms = []
    for p in sorted(tcp_ports)[:max_ports]:
        terms.append(f"({direction_kw} and not impostor and tcp and localPort == {p})")
    for p in sorted(udp_ports)[:max_ports]:
        terms.append(f"({direction_kw} and not impostor and udp and localPort == {p})")
    return " or ".join(terms)


def fmt_ports(ports):
    if not ports:
        return "-"
    items = sorted(ports)
    if len(items) <= 8:
        return ",".join(str(x) for x in items)
    return f"{','.join(str(x) for x in items[:8])}...(+{len(items)-8})"


SPECIAL_VK_NAMES = {
    win32con.VK_LBUTTON: "LBUTTON",
    win32con.VK_RBUTTON: "RBUTTON",
    win32con.VK_MBUTTON: "MBUTTON",
    win32con.VK_XBUTTON1: "XBUTTON1",
    win32con.VK_XBUTTON2: "XBUTTON2",
    win32con.VK_RETURN: "ENTER",
    win32con.VK_SPACE: "SPACE",
    win32con.VK_ESCAPE: "ESC",
    win32con.VK_SHIFT: "SHIFT",
    win32con.VK_CONTROL: "CTRL",
    win32con.VK_MENU: "ALT",
}


def vk_to_name(vk):
    if vk in SPECIAL_VK_NAMES:
        return SPECIAL_VK_NAMES[vk]
    if 0x30 <= vk <= 0x39 or 0x41 <= vk <= 0x5A:
        return chr(vk)
    if 0x70 <= vk <= 0x87:
        return f"F{vk - 0x6F}"
    return f"VK_{vk}"


def is_vk_down(vk):
    return (win32api.GetAsyncKeyState(int(vk)) & 0x8000) != 0


class WindivertAddrRaw(ctypes.Structure):
    _fields_ = [("_qword", ctypes.c_uint64 * 16)]


class WinDivertBindings:
    LAYER_NETWORK = 0
    SHUTDOWN_RECV = 0x1
    INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value

    def __init__(self, dll_path):
        self.dll_path = Path(dll_path).resolve()
        self._dll = ctypes.WinDLL(str(self.dll_path), use_last_error=True)
        self._open = self._dll.WinDivertOpen
        self._open.argtypes = [ctypes.c_char_p, ctypes.c_uint, ctypes.c_int16, ctypes.c_uint64]
        self._open.restype = ctypes.c_void_p

        self._recv = self._dll.WinDivertRecv
        self._recv.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint, ctypes.POINTER(ctypes.c_uint), ctypes.c_void_p]
        self._recv.restype = ctypes.c_bool

        self._send = self._dll.WinDivertSend
        self._send.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint, ctypes.POINTER(ctypes.c_uint), ctypes.c_void_p]
        self._send.restype = ctypes.c_bool

        self._shutdown = self._dll.WinDivertShutdown
        self._shutdown.argtypes = [ctypes.c_void_p, ctypes.c_uint]
        self._shutdown.restype = ctypes.c_bool

        self._close = self._dll.WinDivertClose
        self._close.argtypes = [ctypes.c_void_p]
        self._close.restype = ctypes.c_bool

    def open(self, filter_expr):
        h = self._open(filter_expr.encode("ascii"), self.LAYER_NETWORK, 0, 0)
        if h in (None, 0, self.INVALID_HANDLE_VALUE):
            return None, ctypes.get_last_error()
        return h, 0

    def recv(self, handle, packet_buf, recv_len, addr_buf):
        ok = self._recv(handle, packet_buf, len(packet_buf), ctypes.byref(recv_len), ctypes.byref(addr_buf))
        if not ok:
            return False, ctypes.get_last_error()
        return True, 0

    def send(self, handle, packet_bytes, addr_bytes):
        n = len(packet_bytes)
        if n <= 0:
            return True, 0
        packet_buf = (ctypes.c_ubyte * n).from_buffer_copy(packet_bytes)
        addr_buf = WindivertAddrRaw()
        ctypes.memmove(ctypes.byref(addr_buf), addr_bytes, min(len(addr_bytes), ctypes.sizeof(addr_buf)))
        sent_len = ctypes.c_uint(0)
        ok = self._send(handle, packet_buf, n, ctypes.byref(sent_len), ctypes.byref(addr_buf))
        if not ok:
            return False, ctypes.get_last_error()
        return True, 0

    def shutdown_recv(self, handle):
        ok = self._shutdown(handle, self.SHUTDOWN_RECV)
        if not ok:
            return False, ctypes.get_last_error()
        return True, 0

    def close(self, handle):
        if not handle:
            return True
        return bool(self._close(handle))

@dataclass
class PendingPacket:
    captured_at: float
    expires_at: float
    replay_at: float
    data: bytes
    addr: bytes


class DirectionSession:
    def __init__(
        self,
        direction_id,
        direction_label,
        cfg,
        session_hold_ms,
        handle,
        windivert,
        max_packet_size,
        on_log,
        on_first_capture,
        on_finished,
    ):
        self.direction_id = direction_id
        self.direction_label = direction_label
        self.cfg = deepcopy(cfg)
        self.session_hold_ms = max(0.0, float(session_hold_ms))
        self.handle = handle
        self.windivert = windivert
        self.max_packet_size = int(max_packet_size)

        self._on_log = on_log
        self._on_first_capture = on_first_capture
        self._on_finished = on_finished

        self._lock = threading.Lock()
        self._stop = False
        self._effect_active = True
        self._recv_done = False
        self._finalized = False
        self._first_capture_reported = False

        self._pending_packets = deque()
        self._recv_thread = None
        self._send_thread = None

        self._stats = {
            "captured": 0,
            "sent": 0,
            "dropped_expired": 0,
            "dropped_replay_loss": 0,
            "dropped_drop_mode": 0,
        }

        self._last_drop_log = 0.0
        self._last_loss_log = 0.0
        self._last_dropmode_log = 0.0

    def start(self):
        if self.cfg.mode == "drop":
            self._recv_thread = threading.Thread(target=self._recv_loop_drop, daemon=True)
            self._recv_thread.start()
            return
        self._recv_thread = threading.Thread(target=self._recv_loop_squeeze, daemon=True)
        self._send_thread = threading.Thread(target=self._send_loop_squeeze, daemon=True)
        self._send_thread.start()
        self._recv_thread.start()

    def stop(self):
        with self._lock:
            self._stop = True
            self._effect_active = False
            h = self.handle
        if h:
            self.windivert.shutdown_recv(h)

    def end_effect(self, reason):
        with self._lock:
            if self._stop or not self._effect_active:
                return
            self._effect_active = False
            h = self.handle

        ok, err = self.windivert.shutdown_recv(h)
        if not ok and err not in (6, 995):
            self._log(f"[WARN][{self.direction_label}] WinDivertShutdown 失敗: {err}")

        if self.cfg.mode == "drop":
            self._log(f"[RESTORE][{self.direction_label}] 已結束丟包模式")
        else:
            if reason == "hold_release":
                self._log(f"[REPLAY][{self.direction_label}] 按住模式放開，開始回放")
            elif reason == "left_click":
                self._log(f"[REPLAY][{self.direction_label}] 左鍵中斷，開始回放")
            else:
                self._log(f"[REPLAY][{self.direction_label}] 效果關閉，開始回放")

    def _log(self, text):
        try:
            self._on_log(text)
        except Exception:
            pass

    def _report_first_capture(self):
        with self._lock:
            if self._first_capture_reported:
                return
            self._first_capture_reported = True
        try:
            self._on_first_capture(self.direction_id)
        except Exception:
            pass

    def _purge_expired_locked(self, now):
        dropped_now = 0
        if self._pending_packets:
            kept = deque()
            for pkt in self._pending_packets:
                if now > pkt.expires_at:
                    dropped_now += 1
                else:
                    kept.append(pkt)
            if dropped_now:
                self._pending_packets = kept

        if dropped_now:
            self._stats["dropped_expired"] += dropped_now
            if now - self._last_drop_log >= 1.0:
                self._log(f"[DROP][{self.direction_label}] 超時丟棄累計 {self._stats['dropped_expired']} 個")
                self._last_drop_log = now

    def _recv_loop_squeeze(self):
        packet_buf = (ctypes.c_ubyte * self.max_packet_size)()
        recv_len = ctypes.c_uint(0)
        addr = WindivertAddrRaw()

        while True:
            with self._lock:
                if self._stop or not self.handle:
                    break
                h = self.handle

            ok, err = self.windivert.recv(h, packet_buf, recv_len, addr)
            if not ok:
                if err in (6, 232, 995):
                    break
                time.sleep(0.005)
                continue

            n = recv_len.value
            if n <= 0:
                continue

            packet = bytes(packet_buf[:n])
            addr_bytes = ctypes.string_at(ctypes.byref(addr), ctypes.sizeof(addr))
            now = time.monotonic()

            with self._lock:
                if self._stop:
                    break
                self._stats["captured"] += 1
                self._purge_expired_locked(now)
                if not self._effect_active:
                    continue
                self._pending_packets.append(
                    PendingPacket(
                        captured_at=now,
                        expires_at=now + (self.session_hold_ms / 1000.0),
                        replay_at=0.0,
                        data=packet,
                        addr=addr_bytes,
                    )
                )

            self._report_first_capture()

        with self._lock:
            self._recv_done = True

    def _send_loop_squeeze(self):
        rng = random.Random()
        jitter_lo, jitter_hi = norm_range(self.cfg.replay_jitter_min_ms, self.cfg.replay_jitter_max_ms)
        jitter_enabled = jitter_hi > 0.0

        bw_lo, bw_hi = norm_range(self.cfg.random_bw_min_kbps, self.cfg.random_bw_max_kbps)
        random_bw = (bw_hi > bw_lo) and (bw_hi > 0.0)
        bw_resample_s = max(0.05, float(self.cfg.random_bw_resample_ms) / 1000.0)

        if random_bw:
            cur_kbps = rng.uniform(max(1.0, bw_lo), bw_hi)
        elif bw_hi > 0.0 and bw_lo == bw_hi:
            cur_kbps = bw_hi
        else:
            cur_kbps = 0.0

        rate_bps = kbps_to_bps(cur_kbps)
        limited = rate_bps > 0.0
        token_cap = max(rate_bps * 2.0, 65535.0) if limited else 0.0
        tokens = token_cap
        last_refill = time.monotonic()
        next_bw_resample = last_refill + bw_resample_s
        loss_prob = clamp_pct(self.cfg.replay_loss_rate_pct) / 100.0

        while True:
            now = time.monotonic()
            if random_bw and now >= next_bw_resample:
                cur_kbps = rng.uniform(max(1.0, bw_lo), bw_hi)
                rate_bps = kbps_to_bps(cur_kbps)
                limited = rate_bps > 0.0
                token_cap = max(rate_bps * 2.0, 65535.0) if limited else 0.0
                tokens = min(tokens, token_cap) if limited else 0.0
                next_bw_resample = now + bw_resample_s

            if limited:
                elapsed = now - last_refill
                if elapsed > 0:
                    tokens = min(token_cap, tokens + elapsed * rate_bps)
                    last_refill = now
            else:
                last_refill = now

            pkt = None
            wait_s = 0.01
            finalize = False

            with self._lock:
                if self._stop:
                    finalize = True
                else:
                    self._purge_expired_locked(now)
                    if self._effect_active:
                        wait_s = 0.01
                    else:
                        if self._pending_packets:
                            head = self._pending_packets[0]
                            if head.replay_at <= 0.0:
                                head.replay_at = now + (rng.uniform(jitter_lo, jitter_hi) / 1000.0 if jitter_enabled else 0.0)
                            if now < head.replay_at:
                                wait_s = max(0.001, min(0.1, head.replay_at - now))
                            else:
                                need = len(head.data)
                                if limited and tokens < need:
                                    wait_s = max(0.001, min(0.1, (need - tokens) / rate_bps)) if rate_bps > 0 else 0.05
                                else:
                                    pkt = self._pending_packets.popleft()
                        elif self._recv_done:
                            finalize = True

            if finalize:
                self._finalize()
                break

            if pkt is not None:
                if loss_prob > 0.0 and rng.random() < loss_prob:
                    now_loss = time.monotonic()
                    with self._lock:
                        self._stats["dropped_replay_loss"] += 1
                        if now_loss - self._last_loss_log >= 1.0:
                            self._log(f"[LOSS][{self.direction_label}] 回放丟包累計 {self._stats['dropped_replay_loss']} 個")
                            self._last_loss_log = now_loss
                    continue

                if limited:
                    tokens = max(0.0, tokens - len(pkt.data))

                ok, err = self.windivert.send(self.handle, pkt.data, pkt.addr)
                if not ok:
                    if err in (6, 232, 995):
                        self._finalize()
                        break
                    time.sleep(0.005)
                    continue

                with self._lock:
                    self._stats["sent"] += 1
                continue

            time.sleep(wait_s)

    def _recv_loop_drop(self):
        packet_buf = (ctypes.c_ubyte * self.max_packet_size)()
        recv_len = ctypes.c_uint(0)
        addr = WindivertAddrRaw()
        rng = random.Random()
        drop_prob = clamp_pct(self.cfg.drop_rate_pct) / 100.0

        while True:
            with self._lock:
                if self._stop or not self.handle:
                    break
                h = self.handle

            ok, err = self.windivert.recv(h, packet_buf, recv_len, addr)
            if not ok:
                if err in (6, 232, 995):
                    break
                time.sleep(0.005)
                continue

            n = recv_len.value
            if n <= 0:
                continue
            packet = bytes(packet_buf[:n])
            addr_bytes = ctypes.string_at(ctypes.byref(addr), ctypes.sizeof(addr))
            now = time.monotonic()

            with self._lock:
                if self._stop:
                    break
                self._stats["captured"] += 1
                active = self._effect_active

            self._report_first_capture()

            if active and drop_prob > 0.0 and rng.random() < drop_prob:
                with self._lock:
                    self._stats["dropped_drop_mode"] += 1
                    if now - self._last_dropmode_log >= 1.0:
                        self._log(f"[DROP][{self.direction_label}] 丟包累計 {self._stats['dropped_drop_mode']} 個")
                        self._last_dropmode_log = now
                continue

            ok, err = self.windivert.send(h, packet, addr_bytes)
            if ok:
                with self._lock:
                    self._stats["sent"] += 1
            elif err in (6, 232, 995):
                break
            else:
                time.sleep(0.005)

        with self._lock:
            self._recv_done = True
        self._finalize()

    def _finalize(self):
        with self._lock:
            if self._finalized:
                return
            self._finalized = True
            h = self.handle
            self.handle = None
            self._stop = True
            self._effect_active = False
            self._pending_packets.clear()
            stats = dict(self._stats)

        if h:
            self.windivert.close(h)

        try:
            self._on_finished(self.direction_id, stats)
        except Exception:
            pass

class TrafficEngine(QObject):
    log_signal = Signal(str)
    status_signal = Signal(str)
    effect_signal = Signal(bool)
    busy_signal = Signal(bool)
    hotkey_signal = Signal(str)

    def __init__(self, cfg):
        super().__init__()
        self._lock = threading.Lock()
        self._cfg = deepcopy(cfg)

        self._monitor_running = False
        self._monitor_thread = None
        self._hotkey_enabled = True

        self._sessions = {}
        self._effect_active = False
        self._disconnect_beeped = False

        self._windivert = None
        self._windivert_path = None

    def set_config(self, cfg):
        cfg_copy = deepcopy(cfg)
        cfg_copy.poll_interval_ms = 10
        with self._lock:
            self._cfg = cfg_copy
            name = vk_to_name(self._cfg.trigger_vk)
        self.hotkey_signal.emit(name)

    def get_config_copy(self):
        with self._lock:
            return deepcopy(self._cfg)

    def set_hotkey_enabled(self, enabled):
        with self._lock:
            self._hotkey_enabled = bool(enabled)

    def start_monitoring(self):
        with self._lock:
            if self._monitor_running:
                return
            self._monitor_running = True
        self._monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self._monitor_thread.start()

    def shutdown(self):
        with self._lock:
            self._monitor_running = False
            sessions = list(self._sessions.values())
            self._sessions = {}
            self._effect_active = False
        for s in sessions:
            s.stop()
        if self._monitor_thread and self._monitor_thread.is_alive():
            self._monitor_thread.join(timeout=1.0)

    def manual_toggle_effect(self):
        with self._lock:
            active = self._effect_active
        if active:
            self._end_effect("manual")
        else:
            self._start_effect()

    def manual_end_effect(self):
        self._end_effect("manual")

    def _monitor_loop(self):
        prev_hot = False
        prev_left = False

        while True:
            with self._lock:
                if not self._monitor_running:
                    break
                cfg = deepcopy(self._cfg)
                hotkey_enabled = self._hotkey_enabled
                effect_active = self._effect_active

            poll_s = max(1, int(cfg.poll_interval_ms)) / 1000.0

            if hotkey_enabled:
                hot_down = is_vk_down(cfg.trigger_vk)
                if cfg.trigger_mode == "toggle":
                    if hot_down and not prev_hot:
                        self.manual_toggle_effect()
                else:
                    if hot_down and not effect_active:
                        self._start_effect()
                    elif (not hot_down) and effect_active:
                        self._end_effect("hold_release")
                prev_hot = hot_down

                if cfg.enable_left_click_restore:
                    left_down = is_vk_down(win32con.VK_LBUTTON)
                    with self._lock:
                        still_active = self._effect_active
                    if left_down and not prev_left and still_active:
                        self._end_effect("left_click")
                    prev_left = left_down
                else:
                    prev_left = False
            else:
                prev_hot = False
                prev_left = False

            time.sleep(poll_s)

    def _ensure_windivert(self, cfg):
        dll_path = resolve_windivert_dll_path()
        if not dll_path.exists():
            self.log_signal.emit(f"[ERROR] 找不到 WinDivert DLL: {dll_path}")
            return None

        if self._windivert is None or self._windivert_path != str(dll_path):
            try:
                self._windivert = WinDivertBindings(dll_path)
                self._windivert_path = str(dll_path)
            except OSError as e:
                self.log_signal.emit(f"[ERROR] 載入 WinDivert 失敗: {e}")
                self._windivert = None
                self._windivert_path = None
                return None
        return self._windivert

    def _choose_hold_ms(self, dcfg):
        lo, hi = norm_range(dcfg.random_hold_min_ms, dcfg.random_hold_max_ms)
        if hi > lo:
            return random.uniform(lo, hi)
        return hi

    def _start_effect(self):
        cfg = self.get_config_copy()

        if not is_admin():
            self.log_signal.emit("[ERROR] 請用系統管理員身分執行")
            return

        target = Path(cfg.target_exe)
        if not target.exists():
            self.log_signal.emit(f"[ERROR] 找不到目標程式: {target}")
            return

        with self._lock:
            if self._effect_active:
                return
            if self._sessions:
                self.log_signal.emit("[INFO] 目前仍在回放/收尾中，請稍候")
                return

        windivert = self._ensure_windivert(cfg)
        if windivert is None:
            return

        pids, err_text = get_target_pids(str(target))
        if err_text:
            self.log_signal.emit(f"[WARN] tasklist 失敗: {err_text}")
        if not pids:
            self.log_signal.emit("[WARN] 找不到目標進程，無法啟用效果")
            return

        tcp_ports, udp_ports = collect_ports_for_pids(pids)
        if not tcp_ports and not udp_ports:
            self.log_signal.emit("[WARN] 目標進程目前沒有可用的本機埠")
            return

        directions = [
            ("outbound", "上行", "outbound", cfg.outbound),
            ("inbound", "下行", "inbound", cfg.inbound),
        ]

        sessions = {}
        for dir_id, label, kw, dcfg in directions:
            if not dcfg.enabled:
                continue

            filt = build_filter(kw, tcp_ports, udp_ports, max(1, int(cfg.max_filter_ports)))
            if not filt:
                self.log_signal.emit(f"[WARN][{label}] 無可用過濾條件")
                continue

            handle, err = windivert.open(filt)
            if not handle:
                self.log_signal.emit(f"[ERROR][{label}] WinDivertOpen 失敗: {err}")
                if err:
                    try:
                        self.log_signal.emit(f"[ERROR][{label}] {ctypes.FormatError(err).strip()}")
                    except Exception:
                        pass
                continue

            hold_ms = self._choose_hold_ms(dcfg)
            mode_text = "擠壓" if dcfg.mode == "squeeze" else f"丟包({clamp_pct(dcfg.drop_rate_pct):g}%)"
            self.log_signal.emit(
                f"[READY][{label}] 模式={mode_text} | 本次扣押={hold_ms:g}ms | TCP={fmt_ports(tcp_ports)} UDP={fmt_ports(udp_ports)}"
            )

            sessions[dir_id] = DirectionSession(
                direction_id=dir_id,
                direction_label=label,
                cfg=dcfg,
                session_hold_ms=hold_ms,
                handle=handle,
                windivert=windivert,
                max_packet_size=max(512, int(cfg.max_packet_size)),
                on_log=self.log_signal.emit,
                on_first_capture=self._on_first_capture,
                on_finished=self._on_session_finished,
            )

        if not sessions:
            self.log_signal.emit("[ERROR] 沒有任何方向成功啟用")
            return

        with self._lock:
            if self._effect_active or self._sessions:
                for s in sessions.values():
                    s.stop()
                return
            self._sessions = sessions
            self._effect_active = True
            self._disconnect_beeped = False

        for s in sessions.values():
            s.start()

        self.effect_signal.emit(True)
        self.busy_signal.emit(True)
        self.status_signal.emit("效果啟用中")
        self.log_signal.emit("[HOLD] 效果已啟用，首次成功扣押後才播放提示音")

    def _end_effect(self, reason):
        with self._lock:
            if not self._effect_active:
                return
            self._effect_active = False
            sessions = list(self._sessions.values())

        for s in sessions:
            s.end_effect(reason)

        self.effect_signal.emit(False)
        self.status_signal.emit("回放/收尾中")

    def _on_first_capture(self, direction_id):
        cfg = self.get_config_copy()
        should_beep = False
        with self._lock:
            if not self._disconnect_beeped:
                self._disconnect_beeped = True
                should_beep = True

        if should_beep:
            threading.Thread(
                target=safe_beep,
                args=(cfg.beep_frequency, cfg.beep_duration_ms, winsound.MB_ICONEXCLAMATION),
                daemon=True,
            ).start()
            label = "上行" if direction_id == "outbound" else "下行"
            self.log_signal.emit(f"[CAPTURE] 已確認攔截 {label} 封包")

    def _on_session_finished(self, direction_id, stats):
        label = "上行" if direction_id == "outbound" else "下行"
        self.log_signal.emit(
            f"[DONE][{label}] captured={stats['captured']} sent={stats['sent']} "
            f"expired={stats['dropped_expired']} replay_loss={stats['dropped_replay_loss']} drop_mode={stats['dropped_drop_mode']}"
        )

        with self._lock:
            self._sessions.pop(direction_id, None)
            finished_all = len(self._sessions) == 0
            if finished_all:
                self._disconnect_beeped = False

        if not finished_all:
            return

        cfg = self.get_config_copy()
        threading.Thread(
            target=safe_beep,
            args=(cfg.restore_beep_frequency, cfg.restore_beep_duration_ms, winsound.MB_OK),
            daemon=True,
        ).start()

        self.busy_signal.emit(False)
        self.status_signal.emit("待命")
        self.log_signal.emit("[RESTORE] 所有方向已恢復正常")

def make_dspin(min_v, max_v, step, dec=1):
    s = QDoubleSpinBox()
    s.setRange(min_v, max_v)
    s.setDecimals(dec)
    s.setSingleStep(step)
    return s


class DirectionPanel(QGroupBox):
    def __init__(self, title):
        super().__init__(title)
        self.setObjectName("Card")
        self._build()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 8, 10, 10)
        lay.setSpacing(6)

        top = QHBoxLayout()
        self.enabled = QCheckBox("啟用")
        self.mode = QComboBox()
        self.mode.addItem("擠壓回放", "squeeze")
        self.mode.addItem("丟包模式", "drop")
        top.addWidget(self.enabled)
        top.addStretch(1)
        top.addWidget(QLabel("模式"))
        top.addWidget(self.mode)
        lay.addLayout(top)

        cols = QHBoxLayout()
        cols.setSpacing(10)
        left_form = QFormLayout()
        right_form = QFormLayout()
        for fm in (left_form, right_form):
            fm.setHorizontalSpacing(8)
            fm.setVerticalSpacing(4)

        self.rand_hold_min = make_dspin(0, 120000, 100, 1)
        self.rand_hold_max = make_dspin(0, 120000, 100, 1)
        self.jit_min = make_dspin(0, 10000, 1, 1)
        self.jit_max = make_dspin(0, 10000, 1, 1)
        self.rand_bw_min = make_dspin(0, 10_000_000, 10, 1)
        self.rand_bw_max = make_dspin(0, 10_000_000, 10, 1)
        self.bw_resample = make_dspin(10, 60000, 10, 1)
        self.replay_loss = make_dspin(0, 100, 0.1, 2)
        self.drop_rate = make_dspin(0, 100, 0.1, 2)

        left_form.addRow("隨機扣押最小", self.rand_hold_min)
        left_form.addRow("隨機扣押最大", self.rand_hold_max)
        left_form.addRow("抖動最小(ms)", self.jit_min)
        left_form.addRow("抖動最大(ms)", self.jit_max)
        left_form.addRow("回放丟包率(%)", self.replay_loss)

        right_form.addRow("隨機帶寬最小", self.rand_bw_min)
        right_form.addRow("隨機帶寬最大", self.rand_bw_max)
        right_form.addRow("重抽週期(ms)", self.bw_resample)
        right_form.addRow("丟包模式丟包率", self.drop_rate)

        cols.addLayout(left_form, 1)
        cols.addLayout(right_form, 1)
        lay.addLayout(cols)

        self.mode.currentIndexChanged.connect(self._on_mode_changed)
        self._on_mode_changed()

    def _on_mode_changed(self):
        squeeze = self.mode.currentData() == "squeeze"
        for w in [
            self.rand_hold_min,
            self.rand_hold_max,
            self.jit_min,
            self.jit_max,
            self.rand_bw_min,
            self.rand_bw_max,
            self.bw_resample,
            self.replay_loss,
        ]:
            w.setEnabled(squeeze)
        self.drop_rate.setEnabled(not squeeze)

    def bind_change(self, cb):
        self.enabled.stateChanged.connect(cb)
        self.mode.currentIndexChanged.connect(cb)
        for w in [
            self.rand_hold_min,
            self.rand_hold_max,
            self.jit_min,
            self.jit_max,
            self.rand_bw_min,
            self.rand_bw_max,
            self.bw_resample,
            self.replay_loss,
            self.drop_rate,
        ]:
            w.valueChanged.connect(cb)

    def to_cfg(self):
        return DirectionConfig(
            enabled=self.enabled.isChecked(),
            mode=self.mode.currentData(),
            random_hold_min_ms=self.rand_hold_min.value(),
            random_hold_max_ms=self.rand_hold_max.value(),
            replay_jitter_min_ms=self.jit_min.value(),
            replay_jitter_max_ms=self.jit_max.value(),
            random_bw_min_kbps=self.rand_bw_min.value(),
            random_bw_max_kbps=self.rand_bw_max.value(),
            random_bw_resample_ms=self.bw_resample.value(),
            replay_loss_rate_pct=self.replay_loss.value(),
            drop_rate_pct=self.drop_rate.value(),
        )

    def from_cfg(self, c):
        self.enabled.setChecked(c.enabled)
        self.mode.setCurrentIndex(0 if c.mode == "squeeze" else 1)
        self.rand_hold_min.setValue(c.random_hold_min_ms)
        self.rand_hold_max.setValue(c.random_hold_max_ms)
        self.jit_min.setValue(c.replay_jitter_min_ms)
        self.jit_max.setValue(c.replay_jitter_max_ms)
        self.rand_bw_min.setValue(c.random_bw_min_kbps)
        self.rand_bw_max.setValue(c.random_bw_max_kbps)
        self.bw_resample.setValue(c.random_bw_resample_ms)
        self.replay_loss.setValue(c.replay_loss_rate_pct)
        self.drop_rate.setValue(c.drop_rate_pct)
        self._on_mode_changed()


class TitleBar(QFrame):
    def __init__(self, win):
        super().__init__()
        self._win = win
        self._drag_offset = None
        self.setObjectName("TitleBar")
        l = QHBoxLayout(self)
        l.setContentsMargins(10, 6, 8, 6)
        l.setSpacing(6)

        title = QLabel("RAIN-DelayBurst")
        title.setObjectName("TitleText")
        l.addWidget(title)
        l.addStretch(1)

        self.btn_min = QPushButton("—")
        self.btn_close = QPushButton("✕")
        self.btn_min.setObjectName("TitleBtn")
        self.btn_close.setObjectName("TitleBtnClose")
        self.btn_min.setFixedSize(30, 24)
        self.btn_close.setFixedSize(30, 24)
        self.btn_min.clicked.connect(self._win.showMinimized)
        self.btn_close.clicked.connect(self._win.close)
        l.addWidget(self.btn_min)
        l.addWidget(self.btn_close)

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_offset = e.globalPosition().toPoint() - self.window().frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if self._drag_offset is not None and e.buttons() & Qt.LeftButton:
            self.window().move(e.globalPosition().toPoint() - self._drag_offset)

    def mouseReleaseEvent(self, e):
        self._drag_offset = None

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("RAIN-DelayBurst")
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Window)
        self.setMinimumSize(920, 680)
        self.resize(980, 700)

        self._trigger_vk = DEFAULT_TRIGGER_VK
        self._capture_active = False
        self._capture_prev = {}
        self._capture_armed_at = 0.0
        self._capture_deadline = 0.0

        self._build_ui()
        self._build_engine()
        self._bind_signals()

        cfg = AppConfig()
        self.apply_cfg(cfg)
        self.engine.set_config(cfg)
        self.engine.start_monitoring()

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        out = QVBoxLayout(root)
        out.setContentsMargins(8, 8, 8, 8)
        out.setSpacing(0)

        self.root_frame = QFrame()
        self.root_frame.setObjectName("RootFrame")
        out.addWidget(self.root_frame)

        root_l = QVBoxLayout(self.root_frame)
        root_l.setContentsMargins(0, 0, 0, 0)
        root_l.setSpacing(0)

        self.titlebar = TitleBar(self)
        root_l.addWidget(self.titlebar)

        body = QFrame()
        body.setObjectName("BodyFrame")
        root_l.addWidget(body, 1)
        body_l = QVBoxLayout(body)
        body_l.setContentsMargins(10, 10, 10, 10)
        body_l.setSpacing(10)

        top_row = QHBoxLayout()
        top_row.setSpacing(10)
        body_l.addLayout(top_row)
        self._build_base_card(top_row)
        self._build_advanced_card(top_row)
        self._build_status_ops_card(top_row)

        dir_row = QHBoxLayout()
        dir_row.setSpacing(10)
        self.out_panel = DirectionPanel("上行設定")
        self.in_panel = DirectionPanel("下行設定")
        dir_row.addWidget(self.out_panel, 1)
        dir_row.addWidget(self.in_panel, 1)
        body_l.addLayout(dir_row, 1)

        log_box = QGroupBox("執行紀錄")
        log_box.setObjectName("Card")
        log_l = QVBoxLayout(log_box)
        log_l.setContentsMargins(10, 8, 10, 10)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(120)
        log_l.addWidget(self.log)
        body_l.addWidget(log_box, 1)

        self._apply_style()

    def _build_base_card(self, parent):
        base = QGroupBox("基本設定")
        base.setObjectName("Card")
        parent.addWidget(base, 3)
        g = QGridLayout(base)
        g.setContentsMargins(10, 8, 10, 10)
        g.setHorizontalSpacing(8)
        g.setVerticalSpacing(6)
        g.setColumnStretch(1, 1)

        self.target_edit = QLineEdit()
        self.target_edit.setReadOnly(True)
        self.target_edit.setPlaceholderText("選擇目標程式 .exe")
        self.btn_pick_target = QPushButton("選擇目標")
        self.btn_pick_target.setMinimumWidth(96)
        self.btn_pick_target.clicked.connect(self.pick_target)

        self.hotkey_label = QLabel(vk_to_name(self._trigger_vk))
        self.hotkey_label.setObjectName("ValueLabel")
        self.btn_bind_hotkey = QPushButton("綁定熱鍵")
        self.btn_bind_hotkey.setMinimumWidth(96)
        self.btn_bind_hotkey.clicked.connect(self.toggle_capture)

        self.trigger_mode = QComboBox()
        self.trigger_mode.addItem("切換", "toggle")
        self.trigger_mode.addItem("按住", "hold")
        self.left_restore = QCheckBox("效果期間左鍵立即結束")

        g.addWidget(QLabel("目標程式"), 0, 0, 1, 3)
        g.addWidget(self.target_edit, 1, 0, 1, 3)
        g.addWidget(self.btn_pick_target, 2, 2)
        g.addWidget(QLabel("觸發熱鍵"), 3, 0)
        g.addWidget(self.hotkey_label, 3, 1)
        g.addWidget(self.btn_bind_hotkey, 3, 2)
        g.addWidget(QLabel("熱鍵模式"), 4, 0)
        g.addWidget(self.trigger_mode, 4, 1)
        g.addWidget(self.left_restore, 4, 2)

    def _build_advanced_card(self, parent):
        adv = QGroupBox("進階設定")
        adv.setObjectName("Card")
        parent.addWidget(adv, 2)
        f = QFormLayout(adv)
        f.setContentsMargins(10, 8, 10, 10)
        f.setHorizontalSpacing(10)
        f.setVerticalSpacing(6)
        self.max_ports = QSpinBox(); self.max_ports.setRange(1, 5000)
        self.max_packet = QSpinBox(); self.max_packet.setRange(512, 65535)
        self.beep_freq = QSpinBox(); self.beep_freq.setRange(37, 32767)
        self.beep_dur = QSpinBox(); self.beep_dur.setRange(10, 2000)
        self.restore_beep_freq = QSpinBox(); self.restore_beep_freq.setRange(37, 32767)
        self.restore_beep_dur = QSpinBox(); self.restore_beep_dur.setRange(10, 2000)
        f.addRow("最大過濾埠數", self.max_ports)
        f.addRow("最大封包大小", self.max_packet)
        f.addRow("啟用提示音頻率", self.beep_freq)
        f.addRow("啟用提示音長度", self.beep_dur)
        f.addRow("恢復提示音頻率", self.restore_beep_freq)
        f.addRow("恢復提示音長度", self.restore_beep_dur)

    def _build_status_ops_card(self, parent):
        wrap = QVBoxLayout()
        wrap.setSpacing(10)
        parent.addLayout(wrap, 2)

        stat = QGroupBox("狀態")
        stat.setObjectName("Card")
        f = QFormLayout(stat)
        f.setContentsMargins(10, 8, 10, 10)
        f.setHorizontalSpacing(10)
        f.setVerticalSpacing(6)
        self.status_label = QLabel("待命")
        self.status_label.setObjectName("StatusIdle")
        self.effect_label = QLabel("否")
        self.effect_label.setObjectName("ValueLabel")
        self.busy_label = QLabel("否")
        self.busy_label.setObjectName("ValueLabel")
        self.admin_label = QLabel("是" if is_admin() else "否")
        self.admin_label.setObjectName("ValueLabel")
        self.contact_label = QLabel(
            'Discord: <a href="https://discord.gg/mdFDmCsd">discord.gg/mdFDmCsd</a> | ID: _forc4_'
        )
        self.contact_label.setObjectName("LinkLabel")
        self.contact_label.setOpenExternalLinks(True)
        self.contact_label.setWordWrap(True)
        f.addRow("系統狀態", self.status_label)
        f.addRow("效果啟用", self.effect_label)
        f.addRow("背景忙碌", self.busy_label)
        f.addRow("管理員模式", self.admin_label)
        f.addRow("作者聯繫", self.contact_label)
        wrap.addWidget(stat)

        ops = QGroupBox("操作")
        ops.setObjectName("Card")
        op_l = QHBoxLayout(ops)
        op_l.setContentsMargins(10, 8, 10, 10)
        op_l.setSpacing(8)
        self.btn_toggle = QPushButton("切換效果")
        self.btn_end = QPushButton("結束效果")
        self.btn_save = QPushButton("保存設定")
        self.btn_load = QPushButton("載入設定")
        for btn in (self.btn_toggle, self.btn_end, self.btn_save, self.btn_load):
            btn.setMinimumWidth(92)
        self.btn_toggle.clicked.connect(self.manual_toggle)
        self.btn_end.clicked.connect(self.manual_end)
        self.btn_save.clicked.connect(self.save_cfg)
        self.btn_load.clicked.connect(self.load_cfg)
        op_l.addWidget(self.btn_toggle)
        op_l.addWidget(self.btn_end)
        op_l.addWidget(self.btn_save)
        op_l.addWidget(self.btn_load)
        wrap.addWidget(ops)

    def _apply_style(self):
        self.setStyleSheet(
            """
            * { font-family: "Segoe UI", "Microsoft JhengHei"; font-size: 12px; color: #d5dae0; }
            QFrame#RootFrame {
                background: #0f1115;
                border: 1px solid #232933;
                border-radius: 14px;
            }
            QFrame#TitleBar {
                background: #13171e;
                border-top-left-radius: 14px;
                border-top-right-radius: 14px;
                border-bottom: 1px solid #242b37;
            }
            QLabel#TitleText {
                font-size: 13px;
                font-weight: 700;
                color: #d9e0ca;
                letter-spacing: 0.4px;
            }
            QPushButton#TitleBtn, QPushButton#TitleBtnClose {
                background: #1d232d;
                border: 1px solid #313949;
                border-radius: 6px;
                color: #c7ced7;
            }
            QPushButton#TitleBtn:hover { background: #273040; }
            QPushButton#TitleBtnClose:hover { background: #4a2b31; border-color: #694149; }
            QFrame#BodyFrame {
                background: #11141a;
                border-bottom-left-radius: 14px;
                border-bottom-right-radius: 14px;
            }
            QGroupBox#Card {
                background: #161a21;
                border: 1px solid #2a313d;
                border-radius: 12px;
                margin-top: 10px;
                padding-top: 7px;
            }
            QGroupBox#Card::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 6px;
                color: #abc881;
                font-weight: 700;
            }
            QPushButton {
                background: #222933;
                border: 1px solid #374050;
                border-radius: 9px;
                padding: 5px 10px;
                font-weight: 600;
                color: #d0d7df;
            }
            QPushButton:hover { background: #2d3746; border-color: #4a5668; }
            QPushButton:pressed { background: #1f2530; border-color: #3e4959; }
            QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox, QTextEdit {
                background: #10141b;
                border: 1px solid #2e3644;
                border-radius: 8px;
                padding: 4px 7px;
                color: #dde2e8;
                selection-background-color: #86a94e;
            }
            QComboBox QAbstractItemView {
                background: #141a22;
                border: 1px solid #313949;
                color: #e0e5eb;
            }
            QTextEdit {
                font-family: "Consolas", "Microsoft JhengHei";
                font-size: 11px;
            }
            QLabel#StatusIdle { color: #b7dc68; font-weight: 700; }
            QLabel#ValueLabel { color: #d4e3bd; font-weight: 700; }
            QLabel#LinkLabel { color: #9ca6b1; font-size: 11px; }
            QLabel#LinkLabel a { color: #b7dc68; text-decoration: none; }
            QLabel#LinkLabel a:hover { color: #d1f088; }
            QCheckBox::indicator {
                width: 14px;
                height: 14px;
                border-radius: 4px;
                border: 1px solid #4b5667;
                background: #11161d;
            }
            QCheckBox::indicator:checked {
                background: #9cca4b;
                border: 1px solid #b6df66;
            }
            """
        )

    def _build_engine(self):
        self.engine = TrafficEngine(AppConfig())
        self._sync_timer = QTimer(self)
        self._sync_timer.setSingleShot(True)
        self._sync_timer.timeout.connect(self._sync_cfg_to_engine)

        self._capture_timer = QTimer(self)
        self._capture_timer.setInterval(20)
        self._capture_timer.timeout.connect(self._poll_hotkey_capture)

        self._ui_loading = False

    def _bind_signals(self):
        self.engine.log_signal.connect(self._on_engine_log)
        self.engine.status_signal.connect(self._on_status_update)
        self.engine.effect_signal.connect(self._on_effect_update)
        self.engine.busy_signal.connect(self._on_busy_update)
        self.engine.hotkey_signal.connect(self._on_hotkey_update)

        self.trigger_mode.currentIndexChanged.connect(self._schedule_sync)
        self.left_restore.stateChanged.connect(self._schedule_sync)
        self.target_edit.textChanged.connect(self._schedule_sync)
        self.max_ports.valueChanged.connect(self._schedule_sync)
        self.max_packet.valueChanged.connect(self._schedule_sync)
        self.beep_freq.valueChanged.connect(self._schedule_sync)
        self.beep_dur.valueChanged.connect(self._schedule_sync)
        self.restore_beep_freq.valueChanged.connect(self._schedule_sync)
        self.restore_beep_dur.valueChanged.connect(self._schedule_sync)

        self.out_panel.bind_change(self._schedule_sync)
        self.in_panel.bind_change(self._schedule_sync)

    def _schedule_sync(self, *_):
        if self._ui_loading:
            return
        self._sync_timer.start(120)

    def _sync_cfg_to_engine(self):
        if self._ui_loading:
            return
        self.engine.set_config(self.collect_cfg())

    def collect_cfg(self):
        mode = self.trigger_mode.currentData()
        if mode not in {"toggle", "hold"}:
            mode = "toggle"

        cfg = AppConfig(
            target_exe=self.target_edit.text().strip(),
            windivert_dll="WinDivert.dll",
            trigger_vk=int(self._trigger_vk),
            trigger_mode=mode,
            enable_left_click_restore=self.left_restore.isChecked(),
            poll_interval_ms=10,
            max_filter_ports=max(1, int(self.max_ports.value())),
            max_packet_size=max(512, int(self.max_packet.value())),
            beep_frequency=max(37, int(self.beep_freq.value())),
            beep_duration_ms=max(10, int(self.beep_dur.value())),
            restore_beep_frequency=max(37, int(self.restore_beep_freq.value())),
            restore_beep_duration_ms=max(10, int(self.restore_beep_dur.value())),
            outbound=self.out_panel.to_cfg(),
            inbound=self.in_panel.to_cfg(),
        )
        return AppConfig.from_dict(asdict(cfg))

    def apply_cfg(self, cfg):
        self._ui_loading = True
        try:
            self.target_edit.setText(cfg.target_exe)
            self._trigger_vk = int(cfg.trigger_vk)
            self.hotkey_label.setText(vk_to_name(self._trigger_vk))

            mode_idx = self.trigger_mode.findData(cfg.trigger_mode)
            if mode_idx < 0:
                mode_idx = 0
            self.trigger_mode.setCurrentIndex(mode_idx)
            self.left_restore.setChecked(bool(cfg.enable_left_click_restore))

            self.max_ports.setValue(max(1, int(cfg.max_filter_ports)))
            self.max_packet.setValue(max(512, int(cfg.max_packet_size)))
            self.beep_freq.setValue(max(37, int(cfg.beep_frequency)))
            self.beep_dur.setValue(max(10, int(cfg.beep_duration_ms)))
            self.restore_beep_freq.setValue(max(37, int(cfg.restore_beep_frequency)))
            self.restore_beep_dur.setValue(max(10, int(cfg.restore_beep_duration_ms)))

            self.out_panel.from_cfg(cfg.outbound)
            self.in_panel.from_cfg(cfg.inbound)
        finally:
            self._ui_loading = False

    def pick_target(self):
        cur = self.target_edit.text().strip()
        base_dir = str(Path(cur).parent) if cur else str(Path.cwd())
        path, _ = QFileDialog.getOpenFileName(
            self,
            "選擇目標程式",
            base_dir,
            "Executable (*.exe);;All Files (*.*)",
        )
        if not path:
            return
        self.target_edit.setText(str(Path(path).resolve()))
        self._sync_cfg_to_engine()
        self._on_engine_log(f"[CFG] 目標程式: {self.target_edit.text().strip()}")

    def toggle_capture(self):
        if self._capture_active:
            self._cancel_capture("[CFG] 已取消熱鍵綁定")
            return

        self._capture_active = True
        self._capture_prev = {vk: is_vk_down(vk) for vk in HOTKEY_SCAN_VKS}
        now = time.monotonic()
        self._capture_armed_at = now + 0.2
        self._capture_deadline = now + 10.0
        self.btn_bind_hotkey.setText("請按下一個按鍵...")
        self.engine.set_hotkey_enabled(False)
        self._capture_timer.start()
        self._on_engine_log("[CFG] 熱鍵綁定中，請按下一個按鍵")

    def _cancel_capture(self, log_text=None):
        self._capture_active = False
        self._capture_prev = {}
        self.btn_bind_hotkey.setText("綁定熱鍵")
        self._capture_timer.stop()
        self.engine.set_hotkey_enabled(True)
        if log_text:
            self._on_engine_log(log_text)

    def _poll_hotkey_capture(self):
        if not self._capture_active:
            self._capture_timer.stop()
            return

        now = time.monotonic()
        if now >= self._capture_deadline:
            self._cancel_capture("[WARN] 熱鍵綁定逾時")
            return

        if now < self._capture_armed_at:
            self._capture_prev = {vk: is_vk_down(vk) for vk in HOTKEY_SCAN_VKS}
            return

        for vk in HOTKEY_SCAN_VKS:
            down = is_vk_down(vk)
            prev = self._capture_prev.get(vk, False)
            self._capture_prev[vk] = down
            if down and not prev:
                self._trigger_vk = int(vk)
                name = vk_to_name(self._trigger_vk)
                self.hotkey_label.setText(name)
                self._cancel_capture(f"[CFG] 已綁定熱鍵: {name}")
                self._sync_cfg_to_engine()
                return

    def manual_toggle(self):
        self.engine.manual_toggle_effect()

    def manual_end(self):
        self.engine.manual_end_effect()

    def save_cfg(self):
        cfg = self.collect_cfg()
        path, _ = QFileDialog.getSaveFileName(
            self,
            "保存設定檔",
            str(Path.cwd() / "dc_config.json"),
            "JSON (*.json);;All Files (*.*)",
        )
        if not path:
            return
        target = Path(path)
        try:
            target.write_text(json.dumps(asdict(cfg), ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as e:
            self._on_engine_log(f"[ERROR] 保存設定失敗: {e}")
            return
        self._on_engine_log(f"[CFG] 設定已保存: {target}")

    def load_cfg(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "載入設定檔",
            str(Path.cwd()),
            "JSON (*.json);;All Files (*.*)",
        )
        if not path:
            return
        target = Path(path)
        try:
            data = json.loads(target.read_text(encoding="utf-8"))
            cfg = AppConfig.from_dict(data)
        except Exception as e:
            self._on_engine_log(f"[ERROR] 載入設定失敗: {e}")
            return

        self.apply_cfg(cfg)
        self.engine.set_config(cfg)
        self._on_engine_log(f"[CFG] 已載入設定: {target}")

    def _on_engine_log(self, text):
        stamp = time.strftime("%H:%M:%S")
        self.log.append(f"[{stamp}] {text}")

    def _on_status_update(self, text):
        self.status_label.setText(text)

    def _on_effect_update(self, active):
        self.effect_label.setText("是" if active else "否")

    def _on_busy_update(self, busy):
        self.busy_label.setText("是" if busy else "否")

    def _on_hotkey_update(self, name):
        self.hotkey_label.setText(name)

    def closeEvent(self, event):
        self._cancel_capture()
        self.engine.shutdown()
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
