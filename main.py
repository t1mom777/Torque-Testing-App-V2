import sys
from PyQt6.QtWidgets import QApplication
from db_handler_local import init_db, insert_default_torque_table_data
from modern_torque_app import ModernTorqueApp

def main():
    # Initialize DuckDB with no constraints, then insert default data
    init_db()
    insert_default_torque_table_data()
    
    app = QApplication(sys.argv)
    window = ModernTorqueApp()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
