import xlwings as xw
import os
import sys
import time
import logging
import pythoncom
import win32com.client
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class XLWingsMonitor:
    def __init__(self, excel_path, watch_folder):
        pythoncom.CoInitialize()
        self.excel_path = os.path.abspath(excel_path)
        self.watch_folder = os.path.abspath(watch_folder)
        self.app = None
        self.wb = None
        self.setup_logging()

    def setup_logging(self):
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('xlwings_monitor.log'),
                logging.StreamHandler()
            ]
        )

    def is_excel_running(self):
        try:
            win32com.client.GetActiveObject("Excel.Application")
            return True
        except:
            return False

    def is_workbook_open(self):
        try:
            if not self.app:
                return False
            # Test app connection first
            try:
                _ = self.app.pid
                _ = self.app.api.ActiveWindow  # Add this line
            except:
                self.app = None
                return False
                
            # Then check workbooks
            for wb in self.app.books:
                if os.path.abspath(wb.fullname) == self.excel_path:
                    self.wb = wb
                    return True
            return False
        except:
            return False

    def initialize_excel(self):
        try:
            logging.debug("Début initialisation Excel...")
            pythoncom.CoInitialize()
            
            # Force création d'une nouvelle instance
            self.app = xw.App(visible=True)
            if not self.app:
                raise Exception("Impossible de créer une instance Excel")
                
            logging.debug(f"Nouvelle instance Excel créée: {self.app}")
            
            # Ouvrir le workbook
            logging.debug(f"Tentative d'ouverture du fichier: {self.excel_path}")
            self.wb = self.app.books.open(self.excel_path)
            if not self.wb:
                raise Exception("Impossible d'ouvrir le workbook")
                
            logging.info("Excel initialisé avec succès")
        except Exception as e:
            logging.error(f"Erreur d'initialisation Excel: {str(e)}")
            logging.debug("Stack trace:", exc_info=True)
            raise

    def cleanup(self):
        try:
            if self.wb:
                self.wb.save()
            if self.app:
                self.app.quit()
        except Exception as e:
            logging.error(f"Erreur cleanup: {str(e)}")
        finally:
            pythoncom.CoUninitialize()

    def import_vba_component(self, file_path):
        max_retries = 3
        for attempt in range(max_retries):
            try:
                if not self.is_workbook_open():
                    self.initialize_excel()
                    
                vba = self.wb.api.VBProject
                # Rest of import logic...
                return True
                
            except Exception as e:
                logging.error(f"Attempt {attempt + 1} failed: {str(e)}")
                time.sleep(1)
                
        logging.error(f"Failed to import {file_path} after {max_retries} attempts")

    def start_monitoring(self):
        try:
            self.initialize_excel()
            logging.info(f"Surveillance du dossier: {self.watch_folder}")
            
            event_handler = VBAFileHandler(self)
            observer = Observer()
            observer.schedule(event_handler, self.watch_folder, recursive=False)
            observer.start()
            
            try:
                while True:
                    if not self.is_workbook_open():
                        logging.info("Excel fermé - Arrêt")
                        break
                    time.sleep(1)
            except KeyboardInterrupt:
                pass
            finally:
                observer.stop()
                observer.join()
                
        except Exception as e:
            logging.error(f"Erreur générale: {str(e)}")
        finally:
            self.cleanup()

class VBAFileHandler(FileSystemEventHandler):
    def __init__(self, monitor):
        self.monitor = monitor
        
    def on_modified(self, event):
        if event.is_directory:
            return
        file_path = event.src_path
        if file_path.lower().endswith(('.bas', '.cls', '.frm')):
            self.monitor.import_vba_component(file_path)
            
    def on_created(self, event):
        self.on_modified(event)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python script.py chemin_excel dossier_surveillance")
        sys.exit(1)
    
    monitor = XLWingsMonitor(sys.argv[1], sys.argv[2])
    monitor.start_monitoring()