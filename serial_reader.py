import time

def read_from_serial(port, baud_rate, callback, stop_event):
    """
    Continuously read from the serial port, calling `callback(torque_value)`
    whenever a new torque value is received. This is just a mock that generates data.
    """
    for i in range(5):
        if stop_event.is_set():
            break
        simulated_value = 95 + i
        callback(simulated_value)
        time.sleep(1)

def find_fits_in_selected_row(value, selected_row):
    """
    Checks if 'value' fits within any of the allowance ranges in selected_row.
    Returns a list of dict like [{ 'range_str': 'x-y', 'allowance_index': 1 }, ...].
    """
    results = []
    for i in range(3):
        allow_str = selected_row.get(f"allowance{i+1}", "")
        try:
            low_str, high_str = allow_str.split("-")
            low = float(low_str.strip())
            high = float(high_str.strip())
            if low <= value <= high:
                results.append({
                    "range_str": allow_str,
                    "allowance_index": i + 1
                })
        except:
            pass
    return results
