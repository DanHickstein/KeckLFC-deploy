# -*- coding: utf-8 -*-
import os
import csv
import time
import math
import datetime as dt

from Hardware.SRS_SIM900 import SRS_SIM900, SRS_PIDcontrol_SIM960

# ======== User settings ========
ADDR = 'GPIB0::2::INSTR'
SLOT = 3
NAME = 'Minicomb Intensity Lock Servo'
LOG_DIR = r"C:\Users\KeckLFC\Desktop\Keck\data_test_log"
INTERVAL_SEC = 30 * 60  # 30 minutes
# ===============================


def to_number_or_str(x):
    """Try cast to float; if not, return str(x)."""
    try:
        return float(x)
    except Exception:
        return str(x)


def csv_path_for_today():
    return os.path.join(LOG_DIR, dt.datetime.now().strftime("servo_IM_%Y%m%d.csv"))


def write_row(path, row):
    newfile = not os.path.exists(path)
    with open(path, 'a', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        if newfile or os.stat(path).st_size == 0:
            w.writerow(['timestamp_local', 'output_V', 'input_V', 'note'])
        w.writerow(row)
        f.flush()
        try:
            os.fsync(f.fileno())
        except Exception:
            pass


def connect_servo(addr=ADDR, slot=SLOT, name=NAME, max_tries=3):
    """Connect SIM900 + SIM960 with simple retries (returns srs, servo)."""
    last_err = None
    for attempt in range(1, max_tries + 1):
        try:
            srs = SRS_SIM900(addr=addr)
            srs.connect()
            servo = SRS_PIDcontrol_SIM960(srs, slot, name=name)
            print(f"[INFO] Connected on attempt {attempt}.")
            return srs, servo
        except Exception as e:
            last_err = e
            print(f"[WARN] Connect attempt {attempt}/{max_tries} failed: {e}")
            time.sleep(2 * attempt)
    raise RuntimeError(f"Failed to connect after {max_tries} attempts: {last_err}")


def safe_disconnect(srs):
    """Disconnect/close SRS safely; supports either .disconnect() or .close()."""
    if srs is None:
        return
    try:
        if hasattr(srs, "disconnect"):
            srs.disconnect()
        else:
            srs.close()
    except Exception:
        pass


def measure_once(servo):
    """Return (timestamp, output_V, input_V, note)."""
    ts = dt.datetime.now().isoformat(timespec='seconds')
    out_v = to_number_or_str(servo.get_output_voltage())
    in_v = to_number_or_str(servo.get_measure_input())
    return ts, out_v, in_v, ''


def main():
    os.makedirs(LOG_DIR, exist_ok=True)

    print(f"[INFO] Logging to: {LOG_DIR}")
    print(f"[INFO] Interval: {INTERVAL_SEC//60} minutes. Press Ctrl+C to stop.")

    # 守时调度，避免累计漂移
    next_t = time.time()

    try:
        while True:
            path = csv_path_for_today()

            srs = None
            try:
                # —— 每次到点都临时连接 ——
                srs, servo = connect_servo()
                time.sleep(1)  # 等待稳定
                row = measure_once(servo)
                time.sleep(0.5)  # 等待稳定
            except Exception as e:
                # 连接或测量失败：记录 NaN
                ts = dt.datetime.now().isoformat(timespec='seconds')
                row = (ts, math.nan, math.nan, f'error: {e}')
            finally:
                # —— 立刻断开，在下一次循环再连 —— 
                safe_disconnect(srs)
                time.sleep(0.5)

            write_row(path, row)
            print(f"[INFO] {row[0]}  output={row[1]} V, input={row[2]} V  note='{row[3]}'")

            # 计算下一次触发时间（守时）
            next_t += INTERVAL_SEC
            sleep_s = max(0.0, next_t - time.time())
            time.sleep(sleep_s)

    except KeyboardInterrupt:
        print("\n[INFO] Stopped by user (Ctrl+C).")
    finally:
        print("[INFO] Resources closed.")


if __name__ == "__main__":
    main()
