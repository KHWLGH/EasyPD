"""Device communication helpers for EasyPD."""

from __future__ import annotations

import queue
import time
from typing import Any, Dict, Iterable, Optional

try:
    from witrnhid import WITRN_DEV, is_pdo, is_rdo  # type: ignore
    from witrnhid.core import K2_TARGET_VID, K2_TARGET_PID  # type: ignore
    try:
        import hid  # type: ignore
    except Exception:  # pragma: no cover - backend import guard
        hid = None  # type: ignore
    HID_AVAILABLE = True
except ImportError:  # pragma: no cover - fallback when dependency missing
    HID_AVAILABLE = False
    hid = None  # type: ignore
    print("警告：无法导入 witrnhid 模块，将使用模拟模式")

    class WITRN_DEV:  # type: ignore
        def open(self, *args: Any, **kwargs: Any) -> None:
            raise Exception("HID 库不可用")

        def close(self) -> None:
            pass

        def read_data(self) -> None:
            pass

        def auto_unpack(self):
            return None, None

    def is_pdo(_pkg: Any) -> bool:  # type: ignore
        return False

    def is_rdo(_pkg: Any) -> bool:  # type: ignore
        return False

    K2_TARGET_VID = 0x0000
    K2_TARGET_PID = 0x0000


def enumerate_devices() -> Iterable[Dict[str, Any]]:
    """Return devices matching the default VID/PID pair."""
    if not HID_AVAILABLE or hid is None:
        return []
    return hid.enumerate(K2_TARGET_VID, K2_TARGET_PID)  # type: ignore[arg-type]


def _attempt_open(device: WITRN_DEV, *args: Any, **kwargs: Any) -> bool:
    try:
        device.open(*args, **kwargs)
        return True
    except Exception:
        return False


def _open_with_info(device: WITRN_DEV, info: Optional[Any]) -> bool:
    if info is None:
        return _attempt_open(device)

    if isinstance(info, dict):
        path = info.get("path") or info.get("device_path")
        path_bytes: Optional[bytes] = None
        if isinstance(path, (bytes, bytearray)):
            path_bytes = bytes(path)
        elif isinstance(path, str):
            try:
                path_bytes = path.encode("utf-8", "ignore")
            except Exception:
                path_bytes = None
        if path_bytes and _attempt_open(device, path_bytes):
            return True

        vid = info.get("vid") if info.get("vid") is not None else info.get("vendor_id")
        pid = info.get("pid") if info.get("pid") is not None else info.get("product_id")
        try:
            if vid is not None and pid is not None and _attempt_open(device, int(vid), int(pid)):
                return True
        except Exception:
            pass

        if path and _attempt_open(device, path=path):
            return True
        return _attempt_open(device)

    if isinstance(info, (bytes, bytearray)) and _attempt_open(device, info):
        return True

    if isinstance(info, (list, tuple)) and len(info) == 2:
        try:
            vid_val = int(info[0])
            pid_val = int(info[1])
        except Exception:
            pass
        else:
            if _attempt_open(device, vid_val, pid_val):
                return True

    return _attempt_open(device)


def data_collection_worker(data_queue, stop_event, pause_flag, device_info=None):
    """
    Background process reading PD packets and measurements from the target device.

    This implementation based on witrn_pd_sniffer-3.7.1 design with the following improvements:
    1. Measurements (voltage, current, power) follow the same pause/resume control as PD packets
    2. Measurement refresh rate limited to 5 Hz (5 times per second) for better performance
    3. Uses last_measurement_timestamp to ensure consistent 0.2s interval between updates

    Data extraction uses dictionary-like access: pkg["Current"], pkg["VBus"]
    """

    import struct

    k2 = WITRN_DEV()
    last_measurement_timestamp = 0.0

    try:
        if not _open_with_info(k2, device_info):
            raise Exception("无法打开所选设备")

        while not stop_event.is_set():
            try:
                k2.read_data()
                unpacked = k2.auto_unpack()
                if not unpacked or len(unpacked) != 2:
                    continue

                timestamp_str, pkg = unpacked
                if pkg is None:
                    continue

                payload = {
                    "timestamp": timestamp_str,
                    "time_sec": time.time(),
                    "data": pkg,
                }

                # Try to extract measurements from the package
                measurements = {}

                # Method 1: Try to access via dictionary-like interface (like witrn_pd_sniffer)
                # Based on witrn_pd_sniffer code: pkg["Current"], pkg["VBus"]
                try:
                    current_field = pkg["Current"]
                    if hasattr(current_field, 'value'):
                        current_val = current_field.value()
                        # Ensure current is stored as an absolute value to avoid negative readings
                        try:
                            parsed_current = float(str(current_val).rstrip('A'))
                            measurements['current'] = abs(parsed_current)
                        except Exception:
                            # fallback: try direct conversion without stripping
                            measurements['current'] = abs(float(current_val))
                except Exception:
                    pass

                try:
                    voltage_field = pkg["VBus"]
                    if hasattr(voltage_field, 'value'):
                        voltage_val = voltage_field.value()
                        measurements['voltage'] = float(str(voltage_val).rstrip('V'))
                except Exception:
                    pass

                # Calculate power if we have both
                if 'voltage' in measurements and 'current' in measurements:
                    measurements['power'] = abs(measurements['voltage'] * measurements['current'])

                # Process data based on packet type
                is_pd_packet = False
                try:
                    is_pd_packet = (getattr(pkg, "field", lambda: None)() == "pd")
                except Exception:
                    is_pd_packet = False

                # PD packets
                if is_pd_packet:
                    # Skip PD packets when paused
                    if pause_flag.value == 1:
                        continue
                    try:
                        data_queue.put_nowait(payload)
                    except queue.Full:
                        pass
                # Measurement packets or other data
                elif measurements:
                    # Apply rate limiting: 5 Hz (0.2 seconds interval)
                    should_send_measurement = False
                    if pause_flag.value == 0:  # Only send measurements when not paused
                        now = time.time()
                        if now - last_measurement_timestamp >= 0.2:  # 5 Hz
                            should_send_measurement = True
                            last_measurement_timestamp = now

                    if should_send_measurement:
                        payload["measurements"] = measurements
                        try:
                            data_queue.put_nowait(payload)
                        except queue.Full:
                            pass

            except Exception as exc:
                err_text = str(exc).lower()
                if "read error" in err_text:
                    data_queue.put_nowait({"error": "device_disconnected"})
                    break
                time.sleep(0.01)

    except Exception as exc:
        data_queue.put_nowait({"error": f"connection_failed: {exc}"})
    finally:
        try:
            k2.close()
        except Exception:
            pass


__all__ = [
    "HID_AVAILABLE",
    "K2_TARGET_VID",
    "K2_TARGET_PID",
    "enumerate_devices",
    "data_collection_worker",
    "is_pdo",
    "is_rdo",
]
