from PyQt5 import QtWidgets, uic, QtCore, QtGui
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QUrl, QEvent, QTimer, QSize, QObject, QThread, pyqtSignal, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import QResizeEvent
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
from common.pandas_table import PandasModel
from common.util_functions import GintHandler, AGSHandler
from common.lab_functions import LabHandler
import numpy as np
import sys
import os
import pandas as pd
from configparser import ConfigParser
import webbrowser
import ctypes
from rich import print as rprint
import warnings
warnings.filterwarnings("ignore", category=UserWarning)
pd.options.mode.chained_assignment = None

'window and icon scaling'
QApplication.setHighDpiScaleFactorRoundingPolicy(QtCore.Qt.HighDpiScaleFactorRoundingPolicy.PassThrough) 
QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
appid = 'ags_gui.v.4.5'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(appid)

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        uic.loadUi("common/assets/ui/mainwindow.ui", self)
        self.setWindowIcon(QtGui.QIcon('common/images/geo.ico'))
    
        self.gint_handler = GintHandler()
        self.ags_handler = AGSHandler()
        self.lab_handler = LabHandler()
        self.error_handle = ErrorHandler()
        self.match_thread = ThreadHandler()
        self.player = QMediaPlayer()
        self.config = ConfigParser()
        
        self.config.read('common/assets/settings.ini')
        self.gint_handler.config = self.config
        self.ags_handler.config = self.config
        self.move(200,200)

        self.set_text('''Please insert AGS file.
''')

        'button connects'
        self.button_open.clicked.connect(self.get_ags)
        self.view_data.clicked.connect(self.view_tableview)
        self.button_save_ags.clicked.connect(self.save_ags)
        self.button_count_results.clicked.connect(self.count_lab_results)
        self.button_ags_checker.clicked.connect(self.check_ags)
        self.button_del_tbl.clicked.connect(self.del_non_lab_tables)
        self.button_match_lab.clicked.connect(self.select_lab_match) #Lab selected
        self.button_cpt_only.clicked.connect(self.export_cpt_only)
        self.button_lab_only.clicked.connect(self.export_lab_only)
        self.button_export_results.clicked.connect(self.export_results)
        self.button_export_error.clicked.connect(self.export_errors)
        self.button_convert_excel.clicked.connect(self.convert_excel)

        'table connects'
        self.headings_table.clicked.connect(self.refresh_table)
        self.headings_table.delete_group.connect(lambda x: self.delete_group(x))
        self.headings_table.rename_group.connect(lambda x: self.rename_group(x))
        self.headings_table.new_group.connect(lambda x: self.new_group(x))
        self.tables_table.insert_rows.connect(lambda x: self.add_rows(x))
        self.tables_table.refreshed.connect(self.reload_table)
        self.tables_table.promote_sig.connect(self.promote)

        'handler connects'
        self.gint_handler._disable.connect(self.disable_buttons)
        self.gint_handler._enable.connect(self.enable_buttons)
        self.gint_handler._update_text.connect(self.set_text)
        self.gint_handler._gint_error_flag.connect(lambda x: self.check_gint_error(x))
        self.ags_handler._disable.connect(self.disable_buttons)
        self.ags_handler._enable.connect(self.enable_buttons)
        self.ags_handler._update_text.connect(self.set_text)
        self.ags_handler._open.connect(lambda x: self.button_open.setEnabled(x))
        self.ags_handler._open.connect(lambda x: self.tabWidget.setTabEnabled(1, x))
        self.ags_handler._enable_error_export.connect(lambda x: self.button_export_error.setEnabled(x))
        self.ags_handler._enable_results_export.connect(lambda x: self.button_export_results.setEnabled(x))
        self.ags_handler._set_model.connect(lambda x: self.update_result_model(x))
        self.ags_handler._coin.connect(self.play_coin)
        self.ags_handler._progress_max.connect(lambda x: self.update_progress_max(x))
        self.ags_handler._progress_current.connect(lambda x: self.update_progress_bar(x))
        self.lab_handler._update_text.connect(self.set_text)
        self.lab_handler._nice.connect(self.play_nice)
        self.lab_handler._disable.connect(self.disable_buttons)
        self.lab_handler._enable.connect(self.enable_buttons)
        self.lab_handler._progress_max.connect(lambda x: self.update_progress_max(x))
        self.lab_handler._progress_current.connect(lambda x: self.update_progress_bar(x))
        self.error_handle.err.connect(self.error_handle.show_err)
        self.match_thread.finished.connect(self.lab_match_cleanup)
        
        'set window size'
        self.installEventFilter(self)
        self.set_size()


    def set_text(self, string:str):
        self.text.setText(string)
        QApplication.processEvents()

    def get_gint(self):
        self.gint_handler.get_gint()
    
    def check_gint_error(self, err:bool):
        self.gint_err: bool = None
        self.gint_err = err

    def check_gint(self):
        if len(self.gint_handler.gint_location[0]) == 0:
            self.set_text('''No gINT selected!
AGS file loaded.''')
            return False
        if self.gint_err:
            self.set_text('''64-bit Access Driver not found.
''')
            rprint('[red]64-bit Access Driver not found.[/red]')
            return False
        return True

    def get_ags(self):
        self.progress_bar.setTextVisible(True)
        self.progress_bar.reset()
        self.ags_handler.load_ags_file()
        if len(self.ags_handler.file_location[0]) == 0:
            return
        self.ags_handler.ags_tables_from_file()
        if self.lab_select.currentText() == "Select a Lab":
            self.lab_select.removeItem(0)
            self.lab_select.setCurrentIndex(0)
        self.setup_tables()

    def get_ags_tables(self):
        self.ags_handler.get_ags_tables()

    def get_selected_lab(self):
        return self.lab_select.currentText()

    def select_lab_match(self):
        if self.get_selected_lab() == "Select a Lab" or self.get_selected_lab() == "":
            rprint("[bold]Please selected a Lab to match AGS results to gINT.[bold]")
        elif self.get_selected_lab() == "GM Lab":
            rprint('[purple][bold]GM Lab AGS[/purple][/bold] selected to match to gINT.')
            self.match_unique_id_gqm()
        elif self.get_selected_lab() == "GM Lab PEZ":
            rprint('[purple][bold]GM Lab AGS for PEZ[/purple][/bold] selected to match to gINT.')
            self.match_unique_id_gqm_pez()
        elif self.get_selected_lab() == "DETS":
            rprint('[purple][bold]DETS AGS[/purple][/bold] selected to match to gINT.')
            self.match_unique_id_dets()
        elif self.get_selected_lab() == "DETS PEZ":
            rprint('[purple][bold]DETS AGS for PEZ[/purple][/bold] selected to match to gINT.')
            self.match_unique_id_dets_pez()
        elif self.get_selected_lab() == "Structural Soils":
            rprint('[purple][bold]Structural Soils Soils AGS[/purple][/bold] selected to match to gINT.')
            self.match_unique_id_soils()
        elif self.get_selected_lab() == "Structural Soils PEZ":
            rprint('[purple][bold]Structural Soils Soils AGS for PEZ[/purple][/bold] selected to match to gINT.')
            self.match_unique_id_soils_pez()
        elif self.get_selected_lab() == "PSL":
            rprint('[purple][bold]PSL AGS[/purple][/bold] selected to match to gINT.')
            self.match_unique_id_psl()
        elif self.get_selected_lab() == "Geolabs":
            rprint('[purple][bold]Geolabs AGS[/purple][/bold] selected to match to gINT.')
            self.match_unique_id_geolabs()
        elif self.get_selected_lab() == "Geolabs (50HZ Fugro)":
            print('[purple][bold]Geolabs (50HZ Fugro) AGS[/purple][/bold] selected to match to gINT.')
            self.match_unique_id_geolabs_fugro()
        elif self.get_selected_lab() == "Sinotech TW":
            rprint('[purple][bold]Sinotech (Taiwan) AGS[/purple][/bold] selected to match to gINT.')
            self.match_unique_id_sinotech()
        elif self.get_selected_lab() == "Mewo":
            rprint('[purple][bold]Mewo AGS[/purple][/bold] selected to match to gINT.')
            self.match_unique_id_mewo()

    def update_result_model(self, df):
        model = PandasModel(df)
        self.listbox.setModel(model)
        self.listbox.resizeColumnsToContents()
        self.listbox.horizontalHeader().hide()

    def check_ags(self):
        self.ags_handler.check_ags()

    def export_errors(self):
        self.ags_handler.export_errors()

    def del_non_lab_tables(self):
        self.ags_handler.del_non_lab_tables()
        self.setup_tables()

    def save_ags(self):
        self.ags_handler.save_ags()
    
    def count_lab_results(self):
        self.ags_handler.count_lab_results()
        
    def export_results(self):
        self.ags_handler.export_results()

    def remove_match_id(self):
        self.lab_handler.remove_match_id()


    def setup_tables(self):
        '''setting up the models for groups and tables'''
        table_keys = [k for k in self.ags_handler.tables.keys()]
        table_shapes = [str(f"(x{v.shape[0] - 2})") for k,v in self.ags_handler.tables.items()]
        headings_with_shapes = list(zip(table_keys,table_shapes))
        headings_df = pd.DataFrame.from_dict(headings_with_shapes)
        headings_df.sort_values(0, ascending=True, kind='mergesort', inplace=True, key=lambda col: col.str.lower())
        headings_df.columns = ["",""]
        self._headings_model = PandasModel(headings_df)
        self.headings_table.setModel(self._headings_model)
        self.headings_table.resizeColumnsToContents()
        self.headings_table.horizontalHeader().hide()

        self._tables_model = PandasModel(self.ags_handler.tables[f"{headings_df.iloc[0,0]}"])
        self.tables_table.setModel(self._tables_model)
        self.tables_table.resizeColumnsToContents()
        self.tables_table.horizontalHeader().sectionPressed.connect(self.tables_table.selectColumn)   #col sel

    def refresh_table(self):
        index = self.headings_table.selectionModel().currentIndex()
        self.headings_table.selectRow(index.row())
        value = index.sibling(index.row(),0).data()
        self._tables_model.df = self.ags_handler.tables[value]
        self._tables_model.original = self.ags_handler.tables[value].copy()
        self._tables_model.layoutChanged.emit()
        self.tables_table.resizeColumnsToContents()

    def reload_table(self):
        index = self.headings_table.selectionModel().currentIndex()
        self.headings_table.selectRow(index.row())
        value = index.sibling(index.row(),0).data()
        self.ags_handler.tables[value] = self._tables_model.df
        self._tables_model.original = self.ags_handler.tables[value].copy()
        self._tables_model.layoutChanged.emit()
        self.tables_table.resizeColumnsToContents()

    def view_tableview(self):
        self.tabWidget.setCurrentIndex(1)

    def delete_group(self, group: str):
        try:
            del self.ags_handler.tables[group]
        except Exception as e:
            print(e)
        self.setup_tables()

    def rename_group(self, groups: list):
        try:
            self.ags_handler.tables[groups[1]] = self.ags_handler.tables.pop(groups[0])
        except Exception as e:
            print(e)
        self.setup_tables()

    def new_group(self, group: str):
        temp = {'HEADING': ["UNIT", "TYPE", "DATA"]}
        df = pd.DataFrame(data=temp)
        self.ags_handler.tables[group] = df
        self.setup_tables()

    def add_rows(self, rows: list):
        index = rows[0]
        num_rows = rows[1]
        if self.headings_table.selectionModel().selection().indexes() == []:
            group = self._headings_model.df.iloc[0,0]
        else:
            try:
                selected_rows = self.headings_table.selectionModel().selectedRows()
                group = self._headings_model.data(selected_rows[0], role=QtCore.Qt.DisplayRole)
            except Exception as e:
                print(e)
        num_col = len(self.ags_handler.tables[group].columns)
        try:
            empty_df = pd.DataFrame(data=[[np.nan]*num_col]*num_rows, columns=self.ags_handler.tables[group].columns)
            first_col = list(empty_df.columns)[0]
            empty_df[first_col] = "DATA"
        except Exception as e:
            print(e)
        try:
            new = pd.concat([self.ags_handler.tables[group].loc[:index], empty_df])
            new.replace(np.nan,"", inplace=True)
            new = pd.concat([new, self.ags_handler.tables[group].loc[index+1:]])
            new.reset_index(drop=True, inplace=True)
            self.ags_handler.tables[group] = new
            self.refresh_table()
        except Exception as e:
            print(e)


    def handle_tables(self):
        self.get_ags_tables()
        self.lab_handler.ags_tables = self.ags_handler.ags_tables
        self.lab_handler.tables = self.ags_handler.tables
        self.lab_handler.spec = self.gint_handler.gint_spec

    #===GQM===
    def match_unique_id_gqm(self):
        self.disable_buttons()
        self.get_gint()

        if not self.check_gint():
            return
        
        self.set_text('''Matching Geoquip Lab AGS to gINT,
please wait...''')
        rprint(f"Matching [purple][b]Geoquip Lab[/b][/purple] AGS to gINT... [white][i]{self.gint_handler.gint_location}") 

        self.handle_tables()
        self.match_thread.func = self.lab_handler.match_unique_id_gqm
        self.match_thread.start()
            
    #===DETS===
    def match_unique_id_dets(self):
        self.disable_buttons()
        self.get_gint()

        if not self.check_gint():
            return

        self.set_text('''Matching DETS AGS to gINT,
please wait...''')
        rprint(f"Matching [purple][b]DETS[/purple][/b] AGS to gINT... [white][i]{self.gint_handler.gint_location}") 

        self.handle_tables()
        self.match_thread.func = self.lab_handler.match_unique_id_dets
        self.match_thread.start()


    #===SOILS===
    def match_unique_id_soils(self):
        self.disable_buttons()
        self.get_gint()

        if not self.check_gint():
            return
        
        self.set_text('''Matching Structural Soils AGS to gINT, 
please wait...''')
        rprint(f"Matching [purple][b]Structural Soils[/purple][/b] AGS to gINT... [white][i]{self.gint_handler.gint_location}") 

        self.handle_tables()
        self.match_thread.func = self.lab_handler.match_unique_id_soils
        self.match_thread.start()

    #===PSL===
    def match_unique_id_psl(self):
        self.disable_buttons()
        self.get_gint()

        if not self.check_gint():
            return
        
        self.set_text('''Matching PSL AGS to gINT, 
please wait...''')
        rprint(f"Matching [purple][b]PSL[/purple][b] AGS to gINT... [white][i]{self.gint_handler.gint_location}") 

        self.handle_tables()
        self.match_thread.func = self.lab_handler.match_unique_id_psl
        self.match_thread.start()

    #===GEOLABS===
    def match_unique_id_geolabs(self):
        self.disable_buttons()
        self.get_gint()

        if not self.check_gint():
            return
                
        self.set_text('''Matching Geolabs AGS to gINT, 
please wait...''')
        rprint(f"Matching [purple][b]Geolabs AGS[/purple][/b] to gINT... [white][i]{self.gint_handler.gint_location}") 

        self.handle_tables()
        self.match_thread.func = self.lab_handler.match_unique_id_geolabs
        self.match_thread.start()

    #===GEOLABS FUGRO===
    def match_unique_id_geolabs_fugro(self):
        self.disable_buttons()
        self.get_gint()

        if not self.check_gint():
            return
        
        self.set_text('''Matching Geolabs (Fugro) AGS to gINT, 
please wait...''')
        rprint(f"Matching [purple][b]Geolabs[/purple][/b] AGS to gINT... [white][i]{self.gint_handler.gint_location}") 

        self.handle_tables()
        self.match_thread.func = self.lab_handler.match_unique_id_geolabs_fugro
        self.match_thread.start()

    #===SOILS PEZ===
    def match_unique_id_soils_pez(self):
        self.disable_buttons()
        self.get_gint()

        if not self.check_gint():
            return
        
        self.set_text('''Matching Soils (PEZ) AGS to gINT, 
please wait...''')
        rprint(f"Matching [purple][b]Structural Soils[/purple][/b] AGS to gINT... [white][i]{self.gint_handler.gint_location}") 

        self.handle_tables()
        self.match_thread.func = self.lab_handler.match_unique_id_soils_pez
        self.match_thread.start()

    #===GQM PEZ===
    def match_unique_id_gqm_pez(self):
        self.disable_buttons()
        self.get_gint()

        if not self.check_gint():
            return
        
        self.set_text('''Matching GM Lab (PEZ) AGS to gINT, 
please wait...''')
        rprint(f"Matching [purple][b]GM Lab (PEZ)[/purple][/b] AGS to gINT... [white][i]{self.gint_handler.gint_location}") 

        self.handle_tables()
        self.match_thread.func = self.lab_handler.match_unique_id_gqm_pez
        self.match_thread.start()

    #===DETS PEZ===
    def match_unique_id_dets_pez(self):
        self.disable_buttons()
        self.get_gint()

        if not self.check_gint():
            return
        
        self.set_text('''Matching DETS (PEZ) AGS to gINT, 
please wait...''')
        rprint(f"Matching [purple][b]DETS for PEZ[/purple][/b] AGS to gINT... [white][i]{self.gint_handler.gint_location}") 

        self.handle_tables()
        self.match_thread.func = self.lab_handler.match_unique_id_dets_pez
        self.match_thread.start()

    #===SINOTECH===
    def match_unique_id_sinotech(self):
        self.disable_buttons()
        self.get_gint()

        if not self.check_gint():
            return
        
        self.set_text('''Matching Sinotech AGS to gINT, 
please wait...''')
        rprint(f"Matching [purple][b]Sinotech[/purple][/b] AGS to gINT... [white][i]{self.gint_handler.gint_location}") 

        self.handle_tables()
        self.match_thread.func = self.lab_handler.match_unique_id_sinotech
        self.match_thread.start()


    #===MEWO===
    def match_unique_id_mewo(self):
        self.disable_buttons()
        self.get_gint()

        if not self.check_gint():
            return
        
        self.set_text('''Matching Mewo AGS to gINT, 
please wait...''')
        rprint(f"Matching [purple][b]Mewo[/purple][/b] AGS to gINT... [white][i]{self.gint_handler.gint_location}") 

        self.handle_tables()
        self.match_thread.func = self.lab_handler.match_unique_id_mewo
        self.match_thread.start()


    def lab_match_cleanup(self):
        self.ags_handler.tables = self.lab_handler.tables
        self.remove_match_id()
        self.enable_buttons()
        self.tables_table.resizeColumnsToContents()

    def update_progress_max(self, val):
        self.progress_bar.reset()
        self.progress_bar.setMaximum(val)
        self.progress_bar.setTextVisible(True)
        self.animation = QPropertyAnimation(targetObject=self.progress_bar, propertyName=b"value")
        curve = QEasingCurve()
        curve.setType(QEasingCurve.InOutQuad)
        curve.setAmplitude(0.50)
        curve.setOvershoot(1.70)
        curve.setPeriod(0.50)
        self.animation.setEasingCurve(curve)

    def update_progress_bar(self, val):
        self.animation.setDuration(100)
        self.animation.setStartValue(self.progress_bar.value())
        self.animation.setEndValue(val)
        self.animation.start()


    def export_cpt_only(self):
        self.ags_handler.del_non_cpt_tables()
        self.setup_tables()

    def export_lab_only(self):
        self.ags_handler.export_lab_only()
        self.setup_tables()
    
    def convert_excel(self):
        self.ags_handler.convert_excel()

      
    def play_coin(self):
        coin_num = np.random.randint(100)
        if coin_num == 17:
            coin = ('common/assets/sounds/coin.mp3')
            coin_url = QUrl.fromLocalFile(coin)
            content = QMediaContent(coin_url)
            self.player.setMedia(content)
            self.player.setVolume(22)
            self.player.play()

    def play_nice(self, *args):
        if not args:
            nice_num = np.random.randint(100)
        else:
            try:
                nice_num = np.random.randint(args[0])
            except Exception as e:
                print(e)
        if nice_num == 7:
            nice = ('common/assets/sounds/nice.mp3')
            nice_url = QUrl.fromLocalFile(nice)
            content = QMediaContent(nice_url)
            self.player.setMedia(content)
            self.player.setVolume(33)
            self.player.play()

    def promote(self):
        self.play_nice(8)
        webbrowser.open('https://github.com/lachesis17/AGS-GUI')


    def disable_buttons(self):       
        self.button_open.setEnabled(False)
        self.view_data.setEnabled(False)
        self.button_count_results.setEnabled(False)
        self.button_ags_checker.setEnabled(False)
        self.button_save_ags.setEnabled(False)
        self.button_del_tbl.setEnabled(False)
        self.button_cpt_only.setEnabled(False)
        self.button_lab_only.setEnabled(False)
        self.lab_select.setEnabled(False)
        self.button_match_lab.setEnabled(False)
        self.button_export_results.setEnabled(False)
        self.button_export_error.setEnabled(False)
        self.button_convert_excel.setEnabled(False)
        self.tabWidget.setTabEnabled(1, False)


    def enable_buttons(self):
        self.button_open.setEnabled(True)
        self.view_data.setEnabled(True)
        self.button_count_results.setEnabled(True)
        self.button_ags_checker.setEnabled(True)
        self.button_save_ags.setEnabled(True)
        self.button_del_tbl.setEnabled(True)
        self.button_cpt_only.setEnabled(True)
        self.button_lab_only.setEnabled(True)
        self.lab_select.setEnabled(True)
        self.button_match_lab.setEnabled(True)
        self.button_convert_excel.setEnabled(True)
        self.tabWidget.setTabEnabled(1, True)


    def eventFilter(self, object: QObject, event: QEvent) -> bool:
        if event.type() == QEvent.Type.WindowStateChange:
            maximized = bool(QtCore.Qt.WindowState.WindowMaximized & self.windowState())
            self.config['Window']['maximized'] = str(maximized)
            with open('common/assets/settings.ini', 'w') as configfile: 
                self.config.write(configfile)
        
        return super().eventFilter(object, event)
    
    def resizeEvent(self, event: QResizeEvent) -> None:
        if not self.resizing:
            self.resizing = True
            timer = QTimer()
            timer.singleShot(500,self.on_resize_timer)
            timer.start()
        
        return super().resizeEvent(event)
    
    def on_resize_timer(self):
        if self.isMaximized():
            self.resizing = False
            return
        
        width = str(self.size().width())
        height = str(self.size().height())
        
        self.config['Window']['width'] = width
        self.config['Window']['height'] = height
        with open('common/assets/settings.ini', 'w') as configfile: 
            self.config.write(configfile)
        self.resizing = False
    
    def set_size(self):
        self.resizing = True
        
        if self.config.getboolean('Window','maximized',fallback=False):
            width = self.config['Window']['width']
            height = self.config['Window']['height']
            self.resize(QSize(int(width),int(height)))
            self.showMaximized()
            return
        
        width = self.config['Window']['width']
        height = self.config['Window']['height']
        self.resize(QSize(int(width),int(height)))
        self.resizing = False


class ErrorHandler(QThread):
    err = pyqtSignal(Exception)

    def __init__(self):
        super(ErrorHandler, self).__init__()
        self.func: function
        '''to run a function'''
        # self.error_handle.func = self.lab_handler.match_unique_id_mewo   #error handling test
        # self.error_handle.start()
        # self.error_handle.run_func()

    def run_func(self):
        try:
            return self.func()
        except Exception as e:
            return self.err.emit(e)
        
    def show_err(self, exception):
        print(f'Error: {exception}')
        #self.terminate()


class ThreadHandler(QThread):
    def __init__(self):
        super(ThreadHandler, self).__init__()
        self.func: function

    def run(self):
        try:
            return self.func()
        except Exception as e:
            print(e)
            pass

    def quit(self):
        pass
        
        
def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)

def main():
    app = QtWidgets.QApplication([sys.argv])
    #QtGui.QFontDatabase.addApplicationFont("assets/fonts/Roboto.ttf")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    sys.excepthook = except_hook
    main()
