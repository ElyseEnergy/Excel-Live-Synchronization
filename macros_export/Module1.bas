import xlwings as xw
import os
import sys
import time
import logging
import pythoncom
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
            level=logging.INFO,
            format='%(asctime)s - %(message)s',
            handlers=[
                logging.FileHandler('xlwings_monitor.log'),
                logging.StreamHandler()
            ]
        )

    def initialize_excel(self):
        try:
            # Réinitialiser COM pour chaque opération
            pythoncom.CoInitialize()
            
            # Essayer de se connecter à Excel s'il est déjà ouvert
            try:
                self.app = xw.apps.active
            except:
                self.app = xw.App(visible=True)
                
            if self.app:
                for wb in self.app.books:
                    if os.path.abspath(wb.fullname) == self.excel_path:
                        self.wb = wb
                        break
            
            if not self.wb:
                self.wb = self.app.books.open(self.excel_path)
            
            logging.info("Excel initialisé avec succès via xlwings")
        except Exception as e:
            logging.error(f"Erreur d'initialisation Excel: {str(e)}")
            raise

    def cleanup(self):
        try:
            if self.wb:
                self.wb.save()
        except:
            pass
        finally:
            pythoncom.CoUninitialize()

    def import_vba_component(self, file_path):
        try:
            # Réinitialiser COM pour chaque opération
            pythoncom.CoInitialize()
            
            file_name = os.path.basename(file_path)
            name, ext = os.path.splitext(file_name)
            
            if ext.lower() not in ['.bas', '.cls', '.frm', '.frx']:
                logging.warning(f"Extension non supportée: {ext}")
                return

            # Utiliser l'API VBA via xlwings
            vba = self.wb.api.VBProject
            
            # Supprimer le composant s'il existe
            try:
                existing_comp = vba.VBComponents(name)
                vba.VBComponents.Remove(existing_comp)
            except:
                pass

            # Importer le nouveau composant
            vba.VBComponents.Import(file_path)
            logging.info(f"Composant {name} importé avec succès")
            
        except Exception as e:
            logging.error(f"Erreur lors de l'import de {file_path}: {str(e)}")
        finally:
            pythoncom.CoUninitialize()

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
        
    monitor = XLWingsMonitor(sys.argv[1], sys.argv[2])
    monitor.start_monitoring()