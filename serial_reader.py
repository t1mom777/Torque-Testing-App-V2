import serial
import time
import re

def parse_range(range_str):
    """
    Converts a string like '67.2 - 72.8' into (67.2, 72.8).
    """
    parts = range_str.split('-')
    if len(parts) != 2:
        return None, None
    try:
        low = float(parts[0].strip())
        high = float(parts[1].strip())
        return low, high
    except ValueError:
        return None, None

def within_allowance(target, range_str):
    """
    True if 'target' is within the numeric bounds of 'range_str'.
    Example: target=70.5, range_str='67.2 - 72.8' => True
    """
    low, high = parse_range(range_str)
    if low is None or high is None:
        return False
    return low <= target <= high

def parse_torque_value(line):
    """
    Extracts the first float from the given line of text.
    For example, "HI 301.5 ft.lb" => 301.5
    Returns None if no float is found.
    """
    match = re.search(r"([\d.]+)", line)
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            return None
    return None

def find_fits_in_selected_row(target, row):
    """
    Checks each allowance (allowance1, allowance2, allowance3) in 'row'
    to see if 'target' is in range. Returns a list of matches (dict).
    Each match includes which allowance index and the range string.
    The list is sorted by closeness to the center of the allowance.
    """
    fits = []
    for i in range(1, 4):
        key = f"allowance{i}"
        rng_str = row.get(key, "")
        if within_allowance(target, rng_str):
            low, high = parse_range(rng_str)
            if low is not None and high is not None:
                mid = (low + high) / 2.0
                diff = abs(mid - target)
                fits.append({
                    "row": row,
                    "allowance_index": i,
                    "range_str": rng_str,
                    "diff": diff
                })
    fits.sort(key=lambda x: x["diff"])
    return fits

def read_from_serial(port, baudrate, callback, stop_event=None):
    """
    Continuously reads lines from the specified serial port.
    Each line is parsed for a float. If found, we call callback(float_value).
    If 'stop_event' is set, we stop reading.
    """
    try:
        ser = serial.Serial(port, baudrate, timeout=1)
        print(f"[DEBUG] Starting serial read on {port} at {baudrate} baud...")
    except serial.SerialException as e:
        print("[DEBUG] Could not open serial port:", e)
        return

    try:
        while True:
            if stop_event and stop_event.is_set():
                break

            line = ser.readline().decode('utf-8', errors='replace').strip()
            if line:
                torque_value = parse_torque_value(line)
                if torque_value is not None:
                    print(f"[DEBUG] Serial callback received torque: {torque_value}")
                    callback(torque_value)
            time.sleep(0.01)
    except Exception as e:
        print("[DEBUG] Error in serial reading:", e)
    finally:
        ser.close()
        print("[DEBUG] Serial port closed.")
