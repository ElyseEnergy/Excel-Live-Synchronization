import win32com.client
import os
import sys
import time
import logging
import pythoncom
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class LocalFilesMonitor:
    def __init__(self, excel_path, watch_folder):
        self.excel_path = os.path.abspath(excel_path)
        self.watch_folder = os.path.abspath(watch_folder)
        self.excel = None
        self.wb = None
        self.setup_logging()

    def setup_logging(self):
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(message)s',
            handlers=[
                logging.FileHandler('local_files_monitor.log'),
                logging.StreamHandler()
            ]
        )

    def find_workbook(self):
        excel_instances = self.get_running_instances()
        for excel in excel_instances:
            try:
                for wb in excel.Workbooks:
                    if os.path.abspath(wb.FullName) == self.excel_path:
                        return excel, wb
            except:
                continue
        return None, None

    def get_running_instances(self):
        pythoncom.CoInitialize()
        instances = []
        try:
            instance = win32com.client.GetActiveObject("Excel.Application")
            instances.append(instance)
        except:
            pass
        
        try:
            instance = win32com.client.Dispatch("Excel.Application")
            if instance not in instances:
                instances.append(instance)
        except:
            pass
        
        return instances

    def initialize_excel(self):
        # Chercher d'abord une instance existante
        self.excel, self.wb = self.find_workbook()
        
        if not self.wb:
            # Si aucune instance trouvée, en créer une nouvelle
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = True
            self.wb = self.excel.Workbooks.Open(self.excel_path)
            
        logging.info("Excel initialisé avec succès")

    def cleanup(self):
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    def import_vba_component(self, file_path):
        try:
            file_name = os.path.basename(file_path)
            name, ext = os.path.splitext(file_name)
            
            if ext.lower() not in ['.bas', '.cls', '.frm', '.frx']:
                logging.warning(f"Extension non supportée: {ext}")
                return

            # Supprimer le composant s'il existe déjà
            try:
                existing_comp = self.wb.VBProject.VBComponents(name)
                self.wb.VBProject.VBComponents.Remove(existing_comp)
            except:
                pass

            # Importer le nouveau composant
            self.wb.VBProject.VBComponents.Import(file_path)
            logging.info(f"Composant {name} importé avec succès")
            
        except Exception as e:
            logging.error(f"Erreur lors de l'import de {file_path}: {str(e)}")

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
                    time.sleep(1)
            except KeyboardInterrupt:
                observer.stop()
                observer.join()
                
        except Exception as e:
            logging.error(f"Erreur: {e}")
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
            logging.info(f"Modification détectée: {file_path}")
            self.monitor.import_vba_component(file_path)
            
    def on_created(self, event):
        self.on_modified(event)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python script.py chemin_excel dossier_surveillance")
        sys.exit(1)
        
    monitor = LocalFilesMonitor(sys.argv[1], sys.argv[2])
    monitor.start_monitoring()