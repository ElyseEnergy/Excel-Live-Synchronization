import sys
import logging
from powerquery_monitor import PowerQueryMonitor
from vba_monitor import ExcelVBAMonitor  # On garde l'import pour plus tard

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py chemin_vers_fichier.xlsm")
        sys.exit(1)
        
    monitor = PowerQueryMonitor(sys.argv[1])
    monitor.monitor()