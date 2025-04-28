import win32com.client
import os
import time
import sys
import logging
from datetime import datetime

class ExcelVBAMonitor:
    def __init__(self, excel_path):
        self.excel_path = os.path.abspath(excel_path)
        self.export_path = os.path.join(os.path.dirname(self.excel_path), 'macros_export')
        self.excel = None
        self.wb = None
        self.previous_components = {}
        self.last_known_components = {}
        self.setup_logging()
        os.makedirs(self.export_path, exist_ok=True)
        self.cleanup_old_files()

    def setup_logging(self):
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(message)s',
            handlers=[
                logging.FileHandler('macro_monitor.log'),
                logging.StreamHandler()
            ]
        )

    def get_vba_components(self):
        components = {}
        try:
            for comp in self.wb.VBProject.VBComponents:
                # Ignorer les composants de type Worksheet (100) et ThisWorkbook (100)
                if comp.Type == 100:
                    continue
                    
                count = comp.CodeModule.CountOfLines
                code = ''
                if count > 0:
                    code = comp.CodeModule.Lines(1, count)
                components[comp.Name] = {
                    'type': comp.Type,
                    'code': code
                }
                logging.info(f"Composant trouvé: {comp.Name} (Type: {comp.Type})")
        except Exception as e:
            logging.error(f"Erreur VBA: {e}")
            raise
        return components

    def export_component(self, component, comp_type):
        extensions = {1: '.bas', 2: '.cls', 3: '.frm'}
        if comp_type in extensions:
            file_path = os.path.join(self.export_path, f"{component.Name}{extensions[comp_type]}")
            try:
                component.Export(file_path)
                logging.info(f"Composant exporté: {file_path}")
            except Exception as e:
                logging.error(f"Erreur export {component.Name}: {e}")

    def save_components_from_cache(self):
        """Sauvegarde les composants depuis le cache en fichiers texte"""
        for name, data in self.last_known_components.items():
            file_ext = {1: '.bas', 2: '.cls', 3: '.frm'}.get(data['type'], '.txt')
            file_path = os.path.join(self.export_path, f"{name}{file_ext}")
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(data['code'])
                logging.info(f"Composant sauvegardé depuis le cache: {file_path}")
            except Exception as e:
                logging.error(f"Erreur sauvegarde cache {name}: {e}")

    def initialize_excel(self):
        try:
            self.excel = win32com.client.GetActiveObject("Excel.Application")
            self.excel = win32com.client.Dispatch("Excel.Application")
        except:
            self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = True

    def cleanup(self):
        if self.wb:
            try:
                self.wb.Close(False)
                self.excel.Quit()
            except:
                pass

    def handle_component_changes(self, current_components, force_export=False):
        # Vérifier les modifications et suppressions
        for name, data in self.previous_components.items():
            if name not in current_components:
                logging.info(f"Macro supprimée: {name}")
                try:
                    file_ext = {1: '.bas', 2: '.cls', 3: '.frm'}.get(data['type'], '.txt')
                    file_path = os.path.join(self.export_path, f"{name}{file_ext}")
                    if os.path.exists(file_path):
                        os.remove(file_path)
                        logging.info(f"Fichier supprimé: {file_path}")
                        
                        # Si c'est un UserForm, supprimer aussi le .frx
                        if data['type'] == 3:  # Type 3 = UserForm
                            frx_path = os.path.join(self.export_path, f"{name}.frx")
                            if os.path.exists(frx_path):
                                os.remove(frx_path)
                                logging.info(f"Fichier .frx supprimé: {frx_path}")
                except Exception as e:
                    logging.error(f"Erreur lors de la suppression de {name}: {e}")

        # Vérifier les modifications
        for name, data in current_components.items():
            if force_export or (name not in self.previous_components or 
                self.previous_components[name]['code'] != data['code']):
                logging.info(f"Modification détectée pour {name}")
                comp = self.wb.VBProject.VBComponents(name)
                self.export_component(comp, data['type'])

        # Vérifier les modifications
        for name, data in current_components.items():
            if (name not in self.previous_components or 
                self.previous_components[name]['code'] != data['code']):
                logging.info(f"Modification détectée pour {name}")
                comp = self.wb.VBProject.VBComponents(name)
                self.export_component(comp, data['type'])

    def cleanup_old_files(self):
        """Supprime tous les fichiers de surveillance précédents"""
        for file in os.listdir(self.export_path):
            # Ne garder que les fichiers .bas, .cls, .frm et .frx
            if file.endswith(('.txt', '.frm', '.frx')):
                try:
                    os.remove(os.path.join(self.export_path, file))
                    logging.info(f"Ancien fichier supprimé: {file}")
                except Exception as e:
                    logging.error(f"Erreur suppression {file}: {e}")

        try:
            self.initialize_excel()
            logging.info(f"Surveillance du fichier: {self.excel_path}")
            
            self.wb = self.excel.Workbooks.Open(self.excel_path)
            logging.info("Fichier Excel ouvert avec succès")
            
            # Force l'export initial de tous les composants
            current_components = self.get_vba_components()
            self.last_known_components = current_components.copy()
            self.handle_component_changes(current_components, force_export=True)
            self.previous_components = current_components
            
            while True:
                try:
                    # Vérifier si le workbook est encore ouvert
                    try:
                        _ = self.wb.Name
                        current_components = self.get_vba_components()
                        self.last_known_components = current_components.copy()
                    except:
                        logging.info("Excel fermé par l'utilisateur - Sauvegarde depuis le cache...")
                        self.save_components_from_cache()
                        sys.exit(0)
                    
                    self.handle_component_changes(current_components)
                    self.previous_components = current_components
                    time.sleep(10)
                    
                except Exception as e:
                    if str(e).find("RPC_E_CALL_REJECTED") >= 0:
                        logging.info("Excel fermé - Sauvegarde depuis le cache...")
                        self.save_components_from_cache()
                        sys.exit(0)
                    logging.error(f"Erreur pendant la surveillance: {e}")
                    sys.exit(1)
                    
        except Exception as e:
            logging.error(f"Erreur lors de l'ouverture du fichier: {e}")
            sys.exit(1)
        finally:
            self.cleanup()

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py chemin_vers_fichier.xlsm")
        sys.exit(1)
        
    monitor = ExcelVBAMonitor(sys.argv[1])
    monitor.monitor()