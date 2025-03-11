import duckdb
from typing import Optional

def init_db():
    """
    Creates the database tables using DuckDB with NO constraints.
    """
    conn = duckdb.connect("data.duckdb", read_only=False)
    cursor = conn.cursor()
    
    # Existing tables:
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS TorqueTable (
            id INTEGER,
            max_torque REAL,
            unit TEXT,
            type TEXT,
            applied_torq TEXT,
            allowance1 TEXT,
            allowance2 TEXT,
            allowance3 TEXT
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS RawData (
            id INTEGER,
            torque_value REAL,
            torque_table_id INTEGER,
            allowance_label TEXT,
            range_str TEXT
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Summary (
            id INTEGER,
            allowance_range TEXT,
            test_results TEXT
        )
    """)
    
    # NEW: Create a simple settings table for storing key-value pairs.
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS AppSettings (
            setting_key TEXT,
            setting_value TEXT
        )
    """)
    
    conn.commit()
    conn.close()

def insert_default_torque_table_data():
    """
    Inserts default rows into TorqueTable if it is empty.
    We'll also manually generate IDs for them.
    """
    conn = duckdb.connect("data.duckdb", read_only=False)
    cursor = conn.cursor()
    
    cursor.execute("SELECT COUNT(*) FROM TorqueTable")
    count = cursor.fetchone()[0]
    if count == 0:
        # Let's see what the max ID is so far:
        cursor.execute("SELECT COALESCE(MAX(id), 0) FROM TorqueTable")
        start_id = cursor.fetchone()[0]
        
        # Add 2 sample rows
        sample_data = [
            (
                start_id + 1,
                100, "Nm", "Wrench", "[95, 65, 40]",
                "90.0 - 100.0", "60.0 - 70.0", "36.0 - 44.0"
            ),
            (
                start_id + 2,
                200, "Nm", "Torque Multiplier", "[60, 40, 20]",
                "57.6 - 62.4", "38.4 - 41.6", "19.2 - 20.8"
            )
        ]
        cursor.executemany("""
            INSERT INTO TorqueTable
            (id, max_torque, unit, type, applied_torq, allowance1, allowance2, allowance3)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, sample_data)
        conn.commit()
    conn.close()

def get_torque_table():
    """
    Returns a list of dictionaries representing the TorqueTable rows.
    """
    conn = duckdb.connect("data.duckdb", read_only=False)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM TorqueTable")
    rows = cursor.fetchall()
    columns = [desc[0] for desc in cursor.description]
    conn.close()
    
    result = []
    for row in rows:
        result.append(dict(zip(columns, row)))
    return result

def insert_raw_data(target_torque, row_id, allowance_label, range_str):
    """
    Inserts a raw test reading into RawData, with manual ID generation.
    """
    conn = duckdb.connect("data.duckdb", read_only=False)
    cursor = conn.cursor()
    
    cursor.execute("SELECT COALESCE(MAX(id), 0) + 1 FROM RawData")
    new_id = cursor.fetchone()[0]
    
    cursor.execute("""
        INSERT INTO RawData (id, torque_value, torque_table_id, allowance_label, range_str)
        VALUES (?, ?, ?, ?, ?)
    """, (new_id, target_torque, row_id, allowance_label, range_str))
    
    conn.commit()
    conn.close()

def insert_summary(allow_range, actual_numbers):
    """
    Placeholder function for summary data.
    """
    pass

def add_torque_entry(max_torque, unit, type_, applied_torq, allowance1, allowance2, allowance3):
    """
    Inserts a new entry into TorqueTable, with manual ID generation.
    """
    conn = duckdb.connect("data.duckdb", read_only=False)
    cursor = conn.cursor()
    cursor.execute("SELECT COALESCE(MAX(id), 0) + 1 FROM TorqueTable")
    new_id = cursor.fetchone()[0]
    cursor.execute("""
        INSERT INTO TorqueTable (id, max_torque, unit, type, applied_torq, allowance1, allowance2, allowance3)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (new_id, max_torque, unit, type_, applied_torq, allowance1, allowance2, allowance3))
    conn.commit()
    conn.close()

def update_torque_entry(entry_id, max_torque, unit, type_, applied_torq, allowance1, allowance2, allowance3):
    """
    Updates an existing entry in TorqueTable.
    """
    conn = duckdb.connect("data.duckdb", read_only=False)
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE TorqueTable
        SET max_torque = ?, unit = ?, type = ?, applied_torq = ?,
            allowance1 = ?, allowance2 = ?, allowance3 = ?
        WHERE id = ?
    """, (max_torque, unit, type_, applied_torq, allowance1, allowance2, allowance3, entry_id))
    conn.commit()
    conn.close()

def delete_torque_entry(entry_id):
    """
    Deletes an entry from TorqueTable.
    """
    conn = duckdb.connect("data.duckdb", read_only=False)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM TorqueTable WHERE id = ?", (entry_id,))
    conn.commit()
    conn.close()

# ---------------- NEW: get_app_setting and set_app_setting ----------------

def get_app_setting(key: str) -> Optional[str]:
    """
    Retrieve a setting value from AppSettings by key.
    Returns None if not found.
    """
    conn = duckdb.connect("data.duckdb", read_only=False)
    cursor = conn.cursor()
    cursor.execute("SELECT setting_value FROM AppSettings WHERE setting_key = ?", (key,))
    row = cursor.fetchone()
    conn.close()
    if row:
        return row[0]
    return None

def set_app_setting(key: str, value: str):
    """
    Inserts or updates a setting in AppSettings.
    """
    conn = duckdb.connect("data.duckdb", read_only=False)
    cursor = conn.cursor()
    cursor.execute("SELECT 1 FROM AppSettings WHERE setting_key = ?", (key,))
    row = cursor.fetchone()
    if row:
        cursor.execute("UPDATE AppSettings SET setting_value = ? WHERE setting_key = ?", (value, key))
    else:
        cursor.execute("INSERT INTO AppSettings (setting_key, setting_value) VALUES (?, ?)", (key, value))
    conn.commit()
    conn.close()
