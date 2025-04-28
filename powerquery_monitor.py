import win32com.client
import os
import time
import sys
import logging
from datetime import datetime
import shutil
import json

class PowerQueryMonitor:
    def __init__(self, excel_path):
        self.excel_path = os.path.abspath(excel_path)
        self.export_path = os.path.join(os.path.dirname(self.excel_path), 'powerquery_export')
        self.connections_path = os.path.join(self.export_path, 'Connections')
        self.excel = None
        self.wb = None
        self.previous_queries = {}
        self.last_known_queries = {}
        self.setup_logging()
        os.makedirs(self.export_path, exist_ok=True)
        os.makedirs(self.connections_path, exist_ok=True)
        self.cleanup_old_files()

    def setup_logging(self):
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(message)s',
            handlers=[
                logging.FileHandler('powerquery_monitor.log'),
                logging.StreamHandler()
            ]
        )

    def cleanup_old_files(self):
        # Nettoyage des fichiers .m
        for file in os.listdir(self.export_path):
            if file.endswith('.m'):
                try:
                    os.remove(os.path.join(self.export_path, file))
                    logging.info(f"Ancien fichier Power Query supprimé: {file}")
                except Exception as e:
                    logging.error(f"Erreur suppression {file}: {e}")
        
        # Nettoyage du dossier Connections
        if os.path.exists(self.connections_path):
            shutil.rmtree(self.connections_path)
            os.makedirs(self.connections_path)
            logging.info("Dossier Connections nettoyé")

    def get_power_queries(self):
        queries = {}
        try:
            # Accéder aux requêtes Power Query
            workbook_queries = self.wb.Queries
            for query in workbook_queries:
                queries[query.Name] = {
                    'formula': query.Formula,
                }
                logging.info(f"Requête trouvée: {query.Name}")

            # Copier le dossier Connections s'il existe
            excel_dir = os.path.dirname(self.excel_path)
            workbook_name = os.path.splitext(os.path.basename(self.excel_path))[0]
            logging.info(f"Recherche des connexions dans: {excel_dir}")
            logging.info(f"Nom du workbook: {workbook_name}")
            
            # Chercher avec les deux formats possibles
            source_connections = os.path.join(excel_dir, f"{workbook_name}_Connections")
            if not os.path.exists(source_connections):
                source_connections = os.path.join(excel_dir, "Connections")
            
            if os.path.exists(source_connections):
                logging.info(f"Dossier de connexions trouvé: {source_connections}")
                # Copier tous les fichiers .json
                for file in os.listdir(source_connections):
                    if file.endswith('.json'):
                        source_file = os.path.join(source_connections, file)
                        dest_file = os.path.join(self.connections_path, file)
                        shutil.copy2(source_file, dest_file)
                        logging.info(f"Fichier de connexion copié: {file}")
            else:
                logging.warning(f"Aucun dossier de connexions trouvé à: {source_connections}")
            if os.path.exists(source_connections):
                # Copier tous les fichiers .json
                for file in os.listdir(source_connections):
                    if file.endswith('.json'):
                        shutil.copy2(
                            os.path.join(source_connections, file),
                            os.path.join(self.connections_path, file)
                        )
                        logging.info(f"Fichier de connexion copié: {file}")

        except Exception as e:
            logging.error(f"Erreur Power Query: {e}")
            raise
        return queries

    def save_query(self, name, formula):
        file_path = os.path.join(self.export_path, f"{name}.m")
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(formula)
            logging.info(f"Requête sauvegardée: {file_path}")
        except Exception as e:
            logging.error(f"Erreur sauvegarde requête {name}: {e}")

    def save_queries_from_cache(self):
        for name, data in self.last_known_queries.items():
            file_path = os.path.join(self.export_path, f"{name}.m")
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(data['formula'])
                logging.info(f"Requête sauvegardée depuis le cache: {file_path}")
            except Exception as e:
                logging.error(f"Erreur sauvegarde cache {name}: {e}")

    def handle_query_changes(self, current_queries):
        # Vérifier les modifications et suppressions
        for name, data in self.previous_queries.items():
            if name not in current_queries:
                logging.info(f"Requête supprimée: {name}")
                try:
                    file_path = os.path.join(self.export_path, f"{name}.m")
                    if os.path.exists(file_path):
                        os.remove(file_path)
                        logging.info(f"Fichier supprimé: {file_path}")
                except Exception as e:
                    logging.error(f"Erreur lors de la suppression de {name}: {e}")

        # Vérifier les modifications et ajouts
        for name, data in current_queries.items():
            if (name not in self.previous_queries or 
                self.previous_queries[name]['formula'] != data['formula']):
                logging.info(f"Modification détectée pour {name}")
                self.save_query(name, data['formula'])

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

    def monitor(self):
        try:
            self.initialize_excel()
            logging.info(f"Surveillance des requêtes Power Query: {self.excel_path}")
            
            self.wb = self.excel.Workbooks.Open(self.excel_path)
            logging.info("Fichier Excel ouvert avec succès")
            
            # Force l'export initial de toutes les requêtes
            current_queries = self.get_power_queries()
            self.last_known_queries = current_queries.copy()
            self.handle_query_changes(current_queries)
            self.previous_queries = current_queries
            
            while True:
                try:
                    try:
                        _ = self.wb.Name
                        try:
                            current_queries = self.get_power_queries()
                            self.last_known_queries = current_queries.copy()
                        except Exception as e:
                            if "RPC_E_CALL_REJECTED" not in str(e):
                                logging.info(f"Requête en cours de chargement - attente...")
                                time.sleep(10)
                                continue
                            raise
                    except:
                        logging.info("Excel fermé par l'utilisateur - Sauvegarde depuis le cache...")
                        self.save_queries_from_cache()
                        sys.exit(0)
                    
                    self.handle_query_changes(current_queries)
                    self.previous_queries = current_queries
                    time.sleep(10)
                    
                except Exception as e:
                    if str(e).find("RPC_E_CALL_REJECTED") >= 0:
                        logging.info("Excel fermé - Sauvegarde depuis le cache...")
                        self.save_queries_from_cache()
                        sys.exit(0)
                    logging.error(f"Erreur pendant la surveillance: {e}")
                    sys.exit(1)
                    
        except Exception as e:
            logging.error(f"Erreur lors de l'ouverture du fichier: {e}")
            sys.exit(1)
        finally:
            self.cleanup()