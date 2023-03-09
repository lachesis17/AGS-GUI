from PyQt5 import QtWidgets, uic, QtCore, QtGui
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
from python_ags4 import AGS4
from pandasgui import show
import sys
import os
import pandas as pd
import pyodbc
from statistics import mean
import configparser
import warnings
warnings.filterwarnings("ignore")
QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough) 
QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        uic.loadUi("common/assets/ui/mainwindow.ui", self)
        self.setWindowIcon(QtGui.QIcon('common/images/geobig.ico'))
    
        self.player = QMediaPlayer()
        self.config = configparser.ConfigParser()
        self.config.read('common/assets/settings.ini')
        
        self.move(200,200)
        self.cpt_value = ""

        #self.button_copy_actual.setIcon(QtGui.QIcon('assets/images/copy.png'))
        #self.button_copy_avg.setIcon(QtGui.QIcon('assets/images/copy.png'))
        #self.button_copy_actual.setIconSize(QtCore.QSize(25,25))
        #self.button_copy_avg.setIconSize(QtCore.QSize(25,25))

        #self.dark_mode_button.clicked.connect(self.dark_toggle)
        #self.dark_mode_button.setChecked(bool(self.config.get('Theme','dark')))

        #self.file_open.triggered.connect(self.test)   -menubar

        self.text.setText('''Please insert AGS file.
''')

        self.button_open.clicked.connect(self.get_ags_file)
        self.pandas_gui.clicked.connect(self.start_pandasgui)
        self.button_save_ags.clicked.connect(self.save_ags)
        self.button_count_results.clicked.connect(self.count_lab_results)
        self.button_ags_checker.clicked.connect(self.check_ags)
        self.button_del_tbl.clicked.connect(self.del_non_lab_tables)
        self.button_match_lab.clicked.connect(self.select_lab_match)
        self.button_cpt_only.clicked.connect(self.del_non_cpt_tables)
        self.button_lab_only.clicked.connect(self.export_lab_only)
        self.button_export_results.clicked.connect(self.export_results)
        self.button_export_error.clicked.connect(self.export_errors)

        self.temp_file_name = ''
        self.tables = None
        self.headings = None
        self.gui = None
        self.result_list = []
        self.error_list = []
        self.ags_tables = []
        self.results_with_samp_and_type = ""

        self.core_tables = ["TRAN","PROJ","UNIT","ABBR","TYPE","DICT","LOCA"]

        self.result_tables = ['SAMP','SPEC','TRIG','TRIT','LNMC','LDEN','GRAG','GRAT',
        'CONG','CONS','CODG','CODT','LDYN','LLPL','LPDN','LPEN','LRES','LTCH','LTHC',
        'LVAN','LHVN','RELD','SHBG','SHBT','TREG','TRET','DSSG','DSST','IRSG','IRST',
        'GCHM','ERES','RESG','REST','RESV','RESD','TORV','PTST','RPLT','RCAG','RDEN',
        'RUCS', 'RWCO', 'IRSV']
        
        #set window size
        self.installEventFilter(self)
        self.set_size()
        #self.dark_mode()

    def get_ags_file(self):
        self.disable_buttons()
        self.listbox.clear()

        self.text.setText('''Please insert AGS file.
''')
        QApplication.processEvents()

        if not self.config.get('LastFolder','dir') == "":
            self.file_location = QtWidgets.QFileDialog.getOpenFileNames(self,'Please insert AGS file...', self.config.get('LastFolder','dir'), '*.ags')
        else:
            self.file_location = QtWidgets.QFileDialog.getOpenFileNames(self,'Please insert AGS file...', os.getcwd(), '*.ags')
        try:
            self.file_location = self.file_location[0][0]
            last_dir = str(os.path.dirname(self.file_location))
            self.config.set('LastFolder','dir',last_dir)
            with open('common/assets/settings.ini', 'w') as configfile: 
                self.config.write(configfile)
            self.text.setText('''AGS file loaded.
''')
            QApplication.processEvents()
        except Exception as e:
            print(e)
            self.text.setText('''No AGS file selected!
Please select an AGS with "Open File..."''')
            QApplication.processEvents()
            print("No AGS file selected! Please select an AGS with 'Open File...'")
            self.disable_buttons()
            self.button_open.setEnabled(True)
            return
        
        try:
            self.tables, self.headings = AGS4.AGS4_to_dataframe(self.file_location)
        except:
            print("Uh, something went wrong. Was that an AGS file? Send help.")
            self.button_open.setEnabled(True)
        finally:
            print(f"AGS file loaded: {self.file_location}")

            if self.lab_select.currentText() == "Select a Lab":
                self.lab_select.removeItem(0)
                self.lab_select.setCurrentIndex(0)
            self.enable_buttons()

    def count_lab_results(self):
        self.disable_buttons()

        self.results_with_samp_and_type = pd.DataFrame()

        lab_tables = ['TRIG','LNMC','LDEN','GRAT','CONG','LDYN','LLPL','LPDN','LPEN',
        'LRES','LTCH','LVAN','RELD','SHBG','TREG','DSSG','IRSG','PTST','GCHM','RESG',
        'ERES','RCAG','RDEN','RUCS','RPLT','LHVN'
        ]

        all_results = []
        error_tables = []
        self.result_list = []
        self.ags_table_reset()

        for table in lab_tables:
            if table in list(self.tables):
                self.ags_tables.append(table)
                
        for table in lab_tables:
            if table in self.ags_tables:
                try:
                    location = list(self.tables[table]['LOCA_ID'])
                    location.pop(0)
                    location.pop(0)
                    samp_id = list(self.tables[table]['SAMP_ID'])
                    samp_id.pop(0)
                    samp_id.pop(0)
                    samp_ref = list(self.tables[table]['SPEC_REF'])
                    samp_ref.pop(0)
                    samp_ref.pop(0)
                    samp_depth = list(self.tables[table]['SPEC_DPTH'])
                    samp_depth.pop(0)
                    samp_depth.pop(0)
                    test_type = ""

                    lab_field = [col for col in self.tables[table].columns if 'LAB' in col]
                    if not lab_field == []:
                        lab_field = str(lab_field[0])
                    lab = self.tables[table][lab_field].iloc[2:]

                    if 'GCHM' in table:
                        test_type = list(self.tables[table]['GCHM_CODE'])
                        test_type.pop(0)
                        test_type.pop(0)
                        test_type_df = pd.DataFrame.from_dict(test_type)
                    elif 'TRIG' in table:
                        test_type = list(self.tables[table]['TRIG_COND'])
                        test_type.pop(0)
                        test_type.pop(0)
                        test_type_df = pd.DataFrame.from_dict(test_type)
                    elif 'CONG' in table:
                        test_type = list(self.tables[table]['CONG_COND'])
                        test_type.pop(0)
                        test_type.pop(0)
                        test_type_df = pd.DataFrame.from_dict(test_type)
                    elif 'TREG' in table:
                        test_type = list(self.tables[table]['TREG_TYPE'])
                        test_type.pop(0)
                        test_type.pop(0)
                        test_type_df = pd.DataFrame.from_dict(test_type)
                    elif 'ERES' in table:
                        test_type = list(self.tables[table]['ERES_TNAM'])
                        test_type.pop(0)
                        test_type.pop(0)
                        test_type_df = pd.DataFrame.from_dict(test_type)
                    elif 'GRAT'in table:
                        test_type = list(self.tables[table]['GRAT_TYPE'])
                        test_type.pop(0)
                        test_type.pop(0)
                        samp_with_table = list(zip(location,samp_id,samp_ref,samp_depth,test_type))
                        samp_with_table.pop(0)
                        samp_with_table.pop(0)
                        result_table = pd.DataFrame.from_dict(samp_with_table)
                        result_table.drop_duplicates(inplace=True)
                        result_table.columns = ['POINT','ID','REF','DEPTH','TYPE']
                        tt = result_table['TYPE'].to_list()
                        test_type_df = pd.DataFrame.from_dict(tt)
                    if not type(lab) == list and lab.empty:
                        lab = list(samp_id)
                        for x in range(0,len(lab)):
                            lab[x] = ""

                    samples = list(zip(location,samp_id,samp_ref,samp_depth,test_type,lab))
                    table_results = pd.DataFrame.from_dict(samples)
                    if table == 'GRAT':
                        table_results.drop_duplicates(inplace=True)
        
                    if not test_type == "":
                        table_results.loc[-1] = [table,'','','','','']
                        table_results.index = table_results.index + 1
                        table_results.sort_index(inplace=True)
                        num_test = test_type_df.value_counts()
                        test_counts = pd.DataFrame(num_test)
                        head = []
                        val = []
                        for y in test_counts.index.tolist():
                            head.append(y[0])
                        for z in test_counts.values.tolist():
                            val.append(z)
                        count = list(zip(head,val))
                        lab_count_off = [x for x in lab if x == "Offshore"]
                        lab_count_off = len(lab_count_off)
                        lab_count_none = [x for x in lab if x == ""]
                        lab_count_none = len(lab_count_none)
                        lab_count_on = [x for x in lab if not x == "Offshore" and not x == ""]
                        lab_count_on = len(lab_count_on)
                        #print(f"{table} offshore: {lab_count_off} onshore: {lab_count_on} no lab: {lab_count_none}")
                        if not lab_count_none == 0:
                            if "GRAT" in table:
                                valcount = table_results.shape[0] - 1
                                count = [f"{count}, Onshore:{valcount}"]
                            else:
                                count = [f"{count}, Offshore:{lab_count_off}, Onshore:{lab_count_on}, None:{lab_count_none}"]
                        else:
                            if lab_count_off == 0 and not lab_count_on == 0:
                                count = [f"{count}, Onshore:{lab_count_on}"]
                            elif lab_count_on == 0 and not lab_count_off == 0:
                                count = [f"{count}, Offshore:{lab_count_off}"]
                            else:
                                count = [f"{count}, Offshore:{lab_count_off}, Onshore:{lab_count_on}"]
                    else:
                        test_type = list(samp_id)
                        for x in range(0,len(test_type)):
                            test_type[x] = ""
                        count = str(len(samp_id))
                        lab_count_off = [x for x in lab if x == "Offshore"]
                        lab_count_off = len(lab_count_off)
                        lab_count_none = [x for x in lab if x == ""]
                        lab_count_none = len(lab_count_none)
                        lab_count_on = [x for x in lab if not x == "Offshore" and not x == ""]
                        lab_count_on = len(lab_count_on)
                        #print(f"{table} offshore: {lab_count_off} onshore: {lab_count_on} no lab: {lab_count_none}")
                        if not lab_count_none == 0:
                            count = [f"Offshore:{lab_count_off}, Onshore:{lab_count_on}, None:{lab_count_none}"]
                        else:
                            if lab_count_off == 0 and not lab_count_on == 0:
                                count = [f"Onshore:{lab_count_on}"]
                            elif lab_count_on == 0 and not lab_count_off == 0:
                                count = [f"Offshore:{lab_count_off}"]
                            else:
                                count = [f"Offshore:{lab_count_off}, Onshore:{lab_count_on}"]
                        sample = list(zip(location,samp_id,samp_ref,samp_depth,test_type,lab))
                        table_results_2 = pd.DataFrame.from_dict(sample)
                        if table == 'RPLT':
                            table_results_2.drop_duplicates(inplace=True)
                            count = table_results_2.shape[0]
                            lab_count_off = [x for x in lab if x == "Offshore"]
                            lab_count_off = len(lab_count_off)
                            lab_count_none = [x for x in lab if x == ""]
                            lab_count_none = len(lab_count_none)
                            lab_count_on = [x for x in lab if not x == "Offshore" and not x == ""]
                            lab_count_on = len(lab_count_on)
                            #print(f"{table} offshore: {lab_count_off} onshore: {lab_count_on} no lab: {lab_count_none}")
                            if not lab_count_none == 0:
                                count = [f"Offshore:{lab_count_off}, Onshore:{lab_count_on}, None:{lab_count_none}"]
                            else:
                                if lab_count_off == 0 and not lab_count_on == 0:
                                    count = [f"Onshore:{lab_count_on}"]
                                elif lab_count_on == 0 and not lab_count_off == 0:
                                    count = [f"Offshore:{lab_count_off}"]
                                else:
                                    count = [f"Offshore:{lab_count_off}, Onshore:{lab_count_on}"]
                        table_results_2.loc[-1] = [table,'','','','','']
                        table_results_2.index = table_results_2.index + 1
                        table_results_2.sort_index(inplace=True)
                        table_results = pd.concat([table_results, table_results_2])
                    type_list = []
                    type_list.append(str(table))
                    for x in range(0,len(count)):
                        type_list.append(count[x])
                    all_results.append(type_list)
                    #print(str(table) + " - " + str(type_list))

                    self.results_with_samp_and_type = pd.concat([self.results_with_samp_and_type, table_results])

                except Exception as e:
                    error_tables.append(str(e))

        if error_tables != []:
            print(f"Table(s) not found:  {str(error_tables)}")

        self.result_list = pd.DataFrame.from_dict(all_results, orient='columns')

        if self.result_list.empty:
            df_list = ["Error: No laboratory test results found."]
            empty_df = pd.DataFrame.from_dict(df_list)
            self.result_list = empty_df
        result_list = self.result_list.to_string(col_space=30,justify="center",index=None, header=None)
        print(result_list)
        self.listbox.setText(result_list)

        self.button_export_results.setEnabled(True)

        self.text.setText('''Results list ready to export.
''')
        QApplication.processEvents()
        self.enable_buttons()

        
    def export_results(self):
        self.disable_buttons()

        result_list = self.results_with_samp_and_type.copy(deep=True)
        result_list.reset_index(inplace=True)
        result_list.sort_index(inplace=True)

        if len(result_list) == 5:
            result_list.loc[-1] = ['INDX','BH','ID','REF','DEPTH','LAB']
            result_list.index = result_list.index + 1
            result_list.sort_index(inplace=True)
        else:
            result_list.loc[-1] = ['INDX','BH','ID','REF','DEPTH','TYPE','LAB']
            result_list.index = result_list.index + 1
            result_list.sort_index(inplace=True)
        
        if not self.config.get('LastFolder','dir') == "":
            self.path_directory = QtWidgets.QFileDialog.getSaveFileName(self,'Save results list as...', self.config.get('LastFolder','dir'), '*.csv')
        else:
            self.path_directory = QtWidgets.QFileDialog.getSaveFileName(self,'Save results list as...', os.getcwd(), '*.csv')
        try:
            self.path_directory = self.path_directory[0]
            if not self.path_directory == "":
                all_result_count = self.path_directory[:-4] + "_report_table.csv"
                self.result_list.to_csv(all_result_count, index=False, index_label=False, header=None)
                print(f"File saved in:  + {str(all_result_count)}")
                all_result_filename = self.path_directory[:-4] + "_all_results.csv"
                result_list.to_csv(all_result_filename, index=False,  header=None)	
                print(f"File saved in:  + {str(all_result_filename)}")
            self.enable_buttons()
            self.button_export_results.setEnabled(True)
        except:
            self.enable_buttons()
            self.button_export_results.setEnabled(True)
            return


    def start_pandasgui(self):
        self.disable_buttons()

        self.text.setText('''PandasGUI loading, please wait...
Close GUI to resume.''')
        QApplication.processEvents()
        
        try:
            self.gui = show(**self.tables)
            self.gui.finished.connect(self.update_tables)
        except:
            pass

        self.text.setText('''You can now save the edited AGS.
''')

        QApplication.processEvents()

        #need some way to check if pandasgui is open so that on close it calls the rest of the function so you're working with the updated edited tables
        #unfortunately pandasgui is a custom class and doesnt have pyqt methods like exec(), finished.connect(), etc.
        updated_tables = self.gui.get_dataframes()
        self.tables = updated_tables
        for table in self.result_tables:
            if table in list(self.tables):
                self.ags_tables.append(table)
        self.enable_buttons()
        
    def check_ags(self):
        self.disable_buttons()
        self.text.setText('''Checking AGS for errors...
''')
        QApplication.processEvents()

        try:
            if not self.file_location == '':
                errors = AGS4.check_file(self.file_location)
            else:
                if self.file_location == '':
                    self.text.setText('''No AGS file selected!
Please select an AGS with "Open File..."''')
                    QApplication.processEvents()
                    print("No AGS file selected! Please select an AGS with 'Open File...'")
                else:
                    errors = AGS4.check_file(self.file_location)
                    
        except ValueError as e:
            print(f'AGS Checker ended unexpectedly: {e}')
            return
        
        if not errors:
            print("No errors found. Yay.")
            self.text.setText("""AGS file contains no errors!
""")
            QApplication.processEvents()
            return
        
        for rule, items in errors.items():
            if rule == 'Metadata':
                print('Metadata')
                for msg in items:
                    print(f"{msg['line']}: {msg['desc']}")
                continue
                    
            for error in items:
                self.text.setText('''Error(s) found, check output or click 'Export Error Log'.
''')
                QApplication.processEvents()
                print(f"Error in line: {error['line']}, group: {error['group']}, description: {error['desc']}")
                self.error_list.append(f"Error in line: {error['line']}, group: {error['group']}, description: {error['desc']}")

        if errors:
            self.button_export_error.setEnabled(True)
            err_str = '\n'.join(str(x) for x in self.error_list)
            self.listbox.setText(err_str)
            self.text.setText('''Error(s) found, check output or click 'Export Error Log'.
''')
            QApplication.processEvents()
        self.enable_buttons()

    def export_errors(self):
        self.disable_buttons()
        
        if not self.config.get('LastFolder','dir') == "":
            self.log_path = QtWidgets.QFileDialog.getSaveFileName(self,'Save error log as...', self.config.get('LastFolder','dir'), '*.txt')
        else:
            self.log_path = QtWidgets.QFileDialog.getSaveFileName(self,'Save error log as...', os.getcwd(), '*.txt')
        try:
            self.log_path = self.log_path[0]
            with open(self.log_path, "w") as f:
                for item in self.error_list:
                    f.write("%s\n" % item)
            self.enable_buttons()
            self.button_export_error.setEnabled(True)
        except:
            self.enable_buttons()
            self.button_export_error.setEnabled(True)
            return

        print(f"Error log exported to:  + {str(self.log_path)}")
        self.enable_buttons()
        self.button_export_error.setEnabled(True)


    def save_ags(self):
        self.disable_buttons()

        if not self.config.get('LastFolder','dir') == "":
            newFileName = QtWidgets.QFileDialog.getSaveFileName(self,'Save AGS file as...', self.config.get('LastFolder','dir'), '*.ags')
        else:
            newFileName = QtWidgets.QFileDialog.getSaveFileName(self,'Save AGS file as...', os.getcwd(), '*.ags')
        try:
            newFileName = newFileName[0]
            AGS4.dataframe_to_AGS4(self.tables, self.tables, newFileName)
            print('Done.')
            self.text.setText('''AGS saved.
''')
            QApplication.processEvents()
            self.enable_buttons()
        except:
            self.enable_buttons()
            return

    def get_gint(self):
        self.disable_buttons()

        if not self.config.get('LastFolder','dir') == "":
            self.gint_location = QtWidgets.QFileDialog.getOpenFileNames(self,'Open gINT Project', self.config.get('LastFolder','dir'), '*.gpj')
        else:
            self.gint_location = QtWidgets.QFileDialog.getOpenFileNames(self,'Open gINT Project', os.getcwd(), '*.gpj')
        try:
            self.gint_location = self.gint_location[0][0]
        except:
            msgBox = QMessageBox()
            msgBox.setIcon(QMessageBox.Information)
            msgBox.setText("You must select a gINT")
            msgBox.setWindowTitle("No gINT selected")
            msgBox.setStandardButtons(QMessageBox.Ok)
            msgBox.exec()
            self.enable_buttons()
            return

        self.text.setText('''Getting gINT, please wait...
''')
        QApplication.processEvents()

        try:
            conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+self.gint_location+';')
            query = "SELECT * FROM SPEC"
            self.gint_spec = pd.read_sql(query, conn)
        except Exception as e:
            print(e)
            print("Uhh.... either that's the wrong gINT, or something went wrong.")
            return

    def get_spec(self):
            return self.gint_spec

    def get_ags_tables(self):
        self.ags_table_reset()

        for table in self.result_tables:
            if table in list(self.tables):
                self.ags_tables.append(table)
        return self.ags_tables

    def get_selected_lab(self):
        lab = self.lab_select.currentText()
        return lab

    def select_lab_match(self):
        if self.get_selected_lab() == "Select a Lab" or self.get_selected_lab() == "":
            print("Please selected a Lab to match AGS results to gINT.")
        elif self.get_selected_lab() == "GM Lab":
            print('GM Lab AGS selected to match to gINT.')
            self.match_unique_id_gqm()
        elif self.get_selected_lab() == "GM Lab PEZ":
            print('GM Lab AGS for PEZ selected to match to gINT.')
            self.match_unique_id_gqm_pez()
        elif self.get_selected_lab() == "DETS":
            print('DETS AGS selected to match to gINT.')
            self.match_unique_id_dets()
        elif self.get_selected_lab() == "DETS PEZ":
            print('DETS AGS for PEZ selected to match to gINT.')
            self.match_unique_id_dets_pez()
        elif self.get_selected_lab() == "Structural Soils":
            print('Structural Soils Soils AGS selected to match to gINT.')
            self.match_unique_id_soils()
        elif self.get_selected_lab() == "PSL":
            print('PSL AGS selected to match to gINT.')
            self.match_unique_id_psl()
        elif self.get_selected_lab() == "Geolabs":
            print('Geolabs AGS selected to match to gINT.')
            self.match_unique_id_geolabs()
        elif self.get_selected_lab() == "Geolabs (50HZ Fugro)":
            print('Geolabs (50HZ Fugro) AGS selected to match to gINT.')
            self.match_unique_id_geolabs_fugro()

    def create_match_id(self):
        self.get_ags_tables()

        for table in self.ags_tables:
            try:    
                if 'match_id' not in self.get_spec():
                    self.get_spec().insert(len(list(self.get_spec().columns)),'match_id','')
            
                if 'match_id' not in self.tables[table]:
                    self.tables[table].insert(len(self.tables[table].keys()),'match_id','')
            except Exception as e:
                print(e)
                pass

    def remove_match_id(self):
        self.get_ags_tables()

        for table in self.ags_tables:
            if "match_id" in self.tables[table]:
                self.tables[table].drop(['match_id'], axis=1, inplace=True)

    def check_matched_to_gint(self):
        if self.matched:
            self.text.setText('''Matching complete! Click: 'Save AGS file'.
''')
            QApplication.processEvents()
            print("Matching complete!")
            self.enable_buttons()
            if self.error == True:
                self.text.setText('''gINT matches, Lab doesn't.
Re-open the AGS and select correct lab.''')
                QApplication.processEvents()
        else:    
            self.text.setText('''Couldn't match sample data.
Did you select the correct gINT or AGS?''')
            QApplication.processEvents()
            print("Unable to match sample data from gINT.") 
            self.enable_buttons()


    def match_unique_id_gqm(self):
        self.disable_buttons()
        self.get_gint()
        self.matched = False
        self.error = False

        if not self.gint_location or self.gint_location == '':
            self.text.setText('''AGS file loaded.
''')
            QApplication.processEvents()
            return

        self.text.setText('''Matching AGS to gINT, please wait...
''')
        QApplication.processEvents()
        print(f"Matching GM Lab AGS to gINT... {self.gint_location}") 

        self.get_ags_tables()

        if 'GCHM' in self.ags_tables or 'ERES' in self.ags_tables:
            self.error = True
            print("GCHM or ERES table(s) found.")

        self.get_spec()['Depth'] = self.get_spec()['Depth'].map('{:,.2f}'.format)
        self.get_spec()['Depth'] = self.get_spec()['Depth'].astype(str)
        self.get_spec()['match_id'] = self.get_spec()['PointID']
        self.get_spec()['match_id'] += self.get_spec()['SPEC_REF']
        self.get_spec()['match_id'] += self.get_spec()['Depth']

        for table in self.ags_tables:
            try:
                gint_rows = self.get_spec().shape[0]

                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TYPE']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']

                try:
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                self.matched = True

                                if table == 'CONG':
                                    if self.tables[table]['SPEC_REF'][tablerow] == "OED" or self.tables[table]['SPEC_REF'][tablerow] == "OEDR" and self.tables[table]['CONG_TYPE'][tablerow] == '':
                                        self.tables[table]['CONG_TYPE'][tablerow] = self.tables[table]['SPEC_REF'][tablerow]

                                if table == 'SAMP':
                                    self.tables[table]['SAMP_REM'][tablerow] = self.get_spec()['SPEC_REF'][gintrow]

                                self.tables[table]['SAMP_ID'][tablerow] = self.get_spec()['SAMP_ID'][gintrow]
                                self.tables[table]['SAMP_REF'][tablerow] = self.get_spec()['SAMP_REF'][gintrow]
                                self.tables[table]['SAMP_TYPE'][tablerow] = self.get_spec()['SAMP_TYPE'][gintrow]
                                self.tables[table]['SAMP_TOP'][tablerow] = format(self.get_spec()['SAMP_Depth'][gintrow],'.2f')

                                try:
                                    self.tables[table]['SPEC_REF'][tablerow] = self.get_spec()['SPEC_REF'][gintrow]
                                except:
                                    pass

                                try:
                                    self.tables[table]['SPEC_DPTH'][tablerow] = self.get_spec()['Depth'][gintrow]
                                except:
                                    pass

                                for x in self.tables[table].keys():
                                    if "LAB" in x:
                                        self.tables[table][x][tablerow] = "GM Lab"

                except Exception as e:
                    print(str(e))
                    pass

                '''SHBG'''
                if table == 'SHBG':
                    for tablerow in range(2,len(self.tables[table])):
                        if "small" in str(self.tables[table]['SHBG_TYPE'][tablerow].lower()):
                            self.tables[table]['SHBG_REM'][tablerow] += " - " + self.tables[table]['SHBG_TYPE'][tablerow]
                            self.tables[table]['SHBG_TYPE'][tablerow] = "SMALL SBOX"

                
                '''SHBT'''
                if table == 'SHBT':
                    for tablerow in range(2,len(self.tables[table])):
                        if self.tables[table]['SHBT_NORM'][tablerow]:
                            self.tables[table]['SHBT_NORM'][tablerow] = round(float(self.tables[table]['SHBT_NORM'][tablerow]))


                '''LLPL'''
                if table == 'LLPL':
                    if 'Non-Plastic' not in self.tables[table]:
                        self.tables[table].insert(13,'Non-Plastic','')
                    for tablerow in range(2,len(self.tables[table])):
                        if self.tables[table]['LLPL_LL'][tablerow] == '' and self.tables[table]['LLPL_PL'][tablerow] == '' and self.tables[table]['LLPL_PI'][tablerow] == '':
                            self.tables[table]['Non-Plastic'][tablerow] = -1


                '''GRAG'''
                if table == 'GRAG':
                    for tablerow in range(2,len(self.tables[table])):
                        if self.tables['GRAG']['GRAG_SILT'][tablerow] == '' and self.tables['GRAG']['GRAG_CLAY'][tablerow] == '':
                            if self.tables['GRAG']['GRAG_VCRE'][tablerow] == '':
                                self.tables['GRAG']['GRAG_FINE'][tablerow] = format(100 - (float(self.tables['GRAG']['GRAG_GRAV'][tablerow])) - (float(self.tables['GRAG']['GRAG_SAND'][tablerow])),".1f")
                            else:
                                self.tables['GRAG']['GRAG_FINE'][tablerow] = format(100 - (float(self.tables['GRAG']['GRAG_VCRE'][tablerow])) - (float(self.tables['GRAG']['GRAG_GRAV'][tablerow])) - (float(self.tables['GRAG']['GRAG_SAND'][tablerow])),'.1f')
                        else:
                            self.tables['GRAG']['GRAG_FINE'][tablerow] = format((float(self.tables['GRAG']['GRAG_SILT'][tablerow]) + float(self.tables['GRAG']['GRAG_CLAY'][tablerow])),'.1f')


                '''GRAT'''
                if table == 'GRAT':
                    for tablerow in range(2,len(self.tables[table])):
                        if self.tables[table]['GRAT_PERP'][tablerow]:
                            self.tables[table]['GRAT_PERP'][tablerow] = round(float(self.tables[table]['GRAT_PERP'][tablerow]))


                '''TREG'''
                if table == 'TREG':
                    for tablerow in range(2,len(self.tables[table])):
                        if self.tables[table]['TREG_TYPE'][tablerow] == 'CU' and self.tables[table]['TREG_COH'][tablerow] == '0':
                            self.tables[table]['TREG_COH'][tablerow] = ''
                            self.tables[table]['TREG_PHI'][tablerow] = ''
                            self.tables[table]['TREG_COND'][tablerow] = 'UNDISTURBED'
                        if self.tables[table]['TREG_TYPE'][tablerow] == 'CD':
                            self.tables[table]['TREG_COND'][tablerow] = 'REMOULDED'
                            if self.tables[table]['TREG_PHI'][tablerow] == '':
                                cid_sample = str(self.tables[table]['SAMP_ID'][tablerow]) + "-" + str(self.tables[table]['SPEC_REF'][tablerow])
                                print(f'CID result: {cid_sample} - does not have friction angle.')


                '''TRET'''
                if table == 'TRET':
                    for tablerow in range(2,len(self.tables[table])):
                        if 'TRET_SHST' in self.tables[table].keys():
                            if self.tables[table]['TRET_SHST'][tablerow] == '' and self.tables[table]['TRET_DEVF'][tablerow] != '':
                                if "cell" in str(self.tables['TRET']['TRET_SAT'][tablerow]).lower():
                                    self.tables[table]['TRET_SHST'][tablerow] = round(float(self.tables[table]['TRET_DEVF'][tablerow]) / 2)
                        if 'TRET_CELL' in self.tables[table].keys():
                            if not self.tables[table]['TRET_CELL'][tablerow] == '':
                                self.tables[table]['TRET_CELL'][tablerow] = round(float(self.tables[table]['TRET_CELL'][tablerow]))

                '''LPDN'''
                if table == 'LPDN':
                    for tablerow in range(2,len(self.tables[table])):
                        if self.tables[table]['LPDN_TYPE'][tablerow] == 'LARGE PKY':
                            self.tables[table]['LPDN_TYPE'][tablerow] = 'LARGE PYK'


                '''CONG'''
                if table == 'CONG':
                    for tablerow in range(2,len(self.tables[table])):
                        if self.tables[table]['CONG_TYPE'][tablerow] == '' and self.tables[table]['CONG_COND'][tablerow] == 'Intact':
                            self.tables[table]['CONG_TYPE'][tablerow] = 'CRS'
                            self.tables[table]['CONG_COND'][tablerow] = 'UNDISTURBED'
                        if "intact" in str(self.tables[table]['CONG_COND'][tablerow].lower()):
                            self.tables[table]['CONG_COND'][tablerow] = "UNDISTURBED"
                        if "oed" in str(self.tables[table]['CONG_TYPE'][tablerow].lower()):
                            self.tables[table]['CONG_TYPE'][tablerow] = "IL OEDOMETER"
                            self.tables[table]['CONG_COND'][tablerow] = "UNDISTURBED"
                        self.tables[table]['CONG_COND'][tablerow] = str(self.tables[table]['CONG_COND'][tablerow].upper())


                '''TRIG & TRIT'''
                if table == 'TRIG' or table == 'TRIT':
                    if 'Depth' not in self.tables[table]:
                        self.tables[table].insert(8,'Depth','')
                    if table == 'TRIT':
                        for tablerow in range(2,len(self.tables[table])):
                            if self.tables[table]['TRIT_DEVF'][tablerow]:
                                self.tables[table]['TRIT_DEVF'][tablerow] = round(float(self.tables[table]['TRIT_DEVF'][tablerow]))
                            if self.tables[table]['TRIT_TESN'][tablerow] == '':
                                self.tables[table]['TRIT_TESN'][tablerow] = 1
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                if self.tables['TRIG']['TRIG_COND'][tablerow] == 'REMOULDED':
                                    self.tables[table]['Depth'][tablerow] = float(self.get_spec()['Depth'][gintrow]) + 0.01
                                else:
                                    self.tables[table]['Depth'][tablerow] = self.get_spec()['Depth'][gintrow]


                '''RELD'''
                if table == 'RELD':
                    if 'Depth' not in self.tables[table]:
                        self.tables[table].insert(8,'Depth','')
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                self.tables[table]['Depth'][tablerow] = self.get_spec()['Depth'][gintrow]



                '''RPLT'''
                if table == 'RPLT':
                    if 'Depth' not in self.tables[table]:
                        self.tables[table].insert(8,'Depth','')
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                self.tables[table]['Depth'][tablerow] = self.get_spec()['Depth'][gintrow]
                            if "RPLT_FAIL" in self.tables[table]:
                                if "." in str(self.tables[table]['RPLT_FAIL'][tablerow]):
                                    self.tables[table]['RPLT_FAIL'][tablerow] = float(self.tables[table]['RPLT_FAIL'][tablerow] * 1000) 


                '''RDEN'''
                if table == 'RDEN':
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                if float(self.tables[table]['RDEN_DDEN'][tablerow]) <= 0:
                                    self.tables[table]['RDEN_DDEN'][tablerow] = 0
                                    self.tables[table]['RDEN_PORO'][tablerow] = 0


                '''LDYN'''
                if table == 'LDYN':
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                if 'LDYN_SWAV1' in self.tables[table] or 'LDYN_SWAV1SS' in self.tables[table]:
                                    if self.tables[table]['LDYN_SWAV1SS'][tablerow] == "":
                                        if self.tables[table]['LDYN_SWAV5'][tablerow] == "":
                                            self.tables[table]['LDYN_SWAV'][tablerow] = int(mean([int(float(self.tables[table]['LDYN_SWAV1'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV2'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV3'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV4'][tablerow]))
                                            ]))
                                        else:
                                            self.tables[table]['LDYN_SWAV'][tablerow] = int(mean([int(float(self.tables[table]['LDYN_SWAV1'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV2'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV3'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV4'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV5'][tablerow]))
                                            ]))
                                    else:
                                        if self.tables[table]['LDYN_SWAV5SS'][tablerow] == "":
                                            self.tables[table]['LDYN_SWAV'][tablerow] = int(mean([int(float(self.tables[table]['LDYN_SWAV1SS'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV2SS'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV3SS'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV4SS'][tablerow]))
                                            ]))
                                        else:
                                            self.tables[table]['LDYN_SWAV'][tablerow] = int(mean([int(float(self.tables[table]['LDYN_SWAV1SS'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV2SS'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV3SS'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV4SS'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV5SS'][tablerow]))
                                            ]))
                            if self.tables[table]['LDYN_REM'][tablerow] == "":
                                self.tables[table]['LDYN_REM'][tablerow] = "Bender Element"

            except Exception as e:
                print(f"Couldn't find table or field, skipping... {str(e)}")
                pass

        self.remove_match_id()
        self.check_matched_to_gint()
        self.enable_buttons()
            

    def match_unique_id_dets(self):
        self.disable_buttons()
        self.get_gint()
        self.matched = False
        self.error = False

        if not self.gint_location or self.gint_location == '':
            self.text.setText('''AGS file loaded.
''')
            QApplication.processEvents()
            return

        self.text.setText('''Matching AGS to gINT, please wait...
''')
        QApplication.processEvents()
        print(f"Matching DETS AGS to gINT... {self.gint_location}") 

        self.get_ags_tables()

        if 'GCHM' in self.ags_tables or 'ERES' in self.ags_tables:
            pass
        else:
            self.error = True
            print("Cannot find GCHM or ERES - looks like this AGS is from GM Lab.")

        self.get_spec()['Depth'] = self.get_spec()['Depth'].map('{:,.2f}'.format)
        self.get_spec()['Depth'] = self.get_spec()['Depth'].astype(str)
        self.get_spec()['match_id'] = self.get_spec()['PointID']
        self.get_spec()['match_id'] += self.get_spec()['Depth']

        for table in self.ags_tables:
            try:
                gint_rows = self.get_spec().shape[0]

                for row in range (2,len(self.tables[table])):
                    self.tables[table]['match_id'][row] = str(self.tables[table]['LOCA_ID'][row]).rsplit(' ', 2)[0] + str(self.tables[table]['SAMP_TOP'][row])

                try:
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                self.matched = True
                                if table == 'ERES':
                                    if 'ERES_REM' not in self.tables[table].keys():
                                        self.tables[table].insert(len(self.tables[table].keys()),'ERES_REM','')
                                    self.tables[table]['ERES_REM'][tablerow] = self.tables[table]['SPEC_REF'][tablerow]
                                self.tables[table]['LOCA_ID'][tablerow] = self.get_spec()['PointID'][gintrow]
                                self.tables[table]['SAMP_ID'][tablerow] = self.get_spec()['SAMP_ID'][gintrow]
                                self.tables[table]['SAMP_REF'][tablerow] = self.get_spec()['SAMP_REF'][gintrow]
                                self.tables[table]['SAMP_TYPE'][tablerow] = self.get_spec()['SAMP_TYPE'][gintrow]
                                self.tables[table]['SPEC_REF'][tablerow] = self.get_spec()['SPEC_REF'][gintrow]
                                self.tables[table]['SAMP_TOP'][tablerow] = format(self.get_spec()['SAMP_Depth'][gintrow],'.2f')
                                self.tables[table]['SPEC_DPTH'][tablerow] = self.get_spec()['Depth'][gintrow]
                                
                                for x in self.tables[table].keys():
                                    if "LAB" in x:
                                        self.tables[table][x][tablerow] = "DETS"
                except Exception as e:
                    print(e)
                    pass

                '''GCHM'''
                if table == 'GCHM':
                    for tablerow in range(2,len(self.tables[table])):
                        if "ph" in str(self.tables[table]['GCHM_UNIT'][tablerow].lower()):
                            self.tables[table]['GCHM_UNIT'][tablerow] = "-"
                        if "co3" in str(self.tables[table]['GCHM_CODE'][tablerow].lower()):
                            self.tables[table]['GCHM_CODE'][tablerow] = "CACO3"


                '''ERES'''
                if table == 'ERES':
                    for tablerow in range(2,len(self.tables[table])):
                        if "<" in str(self.tables[table]['ERES_RTXT'][tablerow].lower()):
                            self.tables[table]['ERES_RTXT'][tablerow] = str(self.tables[table]['ERES_RTXT'][tablerow]).rsplit(" ", 1)[1]
                        if "solid_21" in str(self.tables[table]['ERES_REM'][tablerow].lower()) or "2:1" in str(self.tables[table]['ERES_NAME'][tablerow].lower()):
                            self.tables[table]['ERES_NAME'][tablerow] = "SOLID_21 WATER EXTRACT"
                        if "solid_wat" in str(self.tables[table]['ERES_REM'][tablerow].lower()):
                            self.tables[table]['ERES_NAME'][tablerow] = "SOLID_11 WATER EXTRACT"
                        if "solid_tot" in str(self.tables[table]['ERES_REM'][tablerow].lower()):
                            self.tables[table]['ERES_NAME'][tablerow] = "SOLID_TOTAL"
                        if "sulph" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()) and "so4" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()) or "sulf" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                            self.tables[table]['ERES_TNAM'][tablerow] = "WS"
                        if "sulph" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()) and "total" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                            self.tables[table]['ERES_TNAM'][tablerow] = "TS"
                        if "caco3" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                            self.tables[table]['ERES_TNAM'][tablerow] = "CACO3"
                        if "co2" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                            self.tables[table]['ERES_TNAM'][tablerow] = "CO2"
                        if "ph" == str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                            self.tables[table]['ERES_TNAM'][tablerow] = "PH"
                        if "chloride" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                            self.tables[table]['ERES_TNAM'][tablerow] = "Cl"
                        if "los" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                            self.tables[table]['ERES_TNAM'][tablerow] = "LOI"
                        if "ph" in str(self.tables[table]['ERES_RUNI'][tablerow].lower()):
                            self.tables[table]['ERES_RUNI'][tablerow] = "-"

            except Exception as e:
                print(f"Couldn't find table or field, skipping... {str(e)}")
                pass

        self.remove_match_id()
        self.check_matched_to_gint()
        self.enable_buttons()

    
    def match_unique_id_soils(self):
        self.disable_buttons()
        self.get_gint()
        self.matched = False
        self.error = False

        if not self.gint_location or self.gint_location == '':
            self.text.setText('''AGS file loaded.
''')
            QApplication.processEvents()
            return

        self.text.setText('''Matching AGS to gINT, please wait...
''')
        QApplication.processEvents()
        print(f"Matching Structural Soils AGS to gINT... {self.gint_location}") 

        self.get_ags_tables()

        self.get_spec()['Depth'] = self.get_spec()['Depth'].map('{:,.2f}'.format)
        self.get_spec()['Depth'] = self.get_spec()['Depth'].astype(str)
        self.get_spec()['match_id'] = self.get_spec()['PointID']
        self.get_spec()['match_id'] += self.get_spec()['Depth']


        for table in self.ags_tables:
            try:
                gint_rows = self.get_spec().shape[0]

                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']

                try:
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                self.matched = True
                                self.tables[table]['LOCA_ID'][tablerow] = self.get_spec()['PointID'][gintrow]
                                self.tables[table]['SAMP_ID'][tablerow] = self.get_spec()['SAMP_ID'][gintrow]
                                self.tables[table]['SAMP_REF'][tablerow] = self.get_spec()['SAMP_REF'][gintrow]
                                self.tables[table]['SAMP_TYPE'][tablerow] = self.get_spec()['SAMP_TYPE'][gintrow]
                                self.tables[table]['SPEC_REF'][tablerow] = self.get_spec()['SPEC_REF'][gintrow]
                                self.tables[table]['SAMP_TOP'][tablerow] = format(self.get_spec()['SAMP_Depth'][gintrow],'.2f')
                                self.tables[table]['SPEC_DPTH'][tablerow] = self.get_spec()['Depth'][gintrow]
                                
                                for x in self.tables[table].keys():
                                    if "LAB" in x:
                                        self.tables[table][x][tablerow] = "Structural Soils"
                except:
                    pass

                '''CONG'''
                if table == 'CONG':
                    for tablerow in range(2,len(self.tables[table])):
                        if "undisturbed" in str(self.tables[table]['CONG_COND'][tablerow].lower()):
                            self.tables[table]['CONG_COND'][tablerow] = "UNDISTURBED"
                        if "oed" in str(self.tables[table]['CONG_TYPE'][tablerow].lower()):
                            self.tables[table]['CONG_TYPE'][tablerow] = "IL OEDOMETER"
                            self.tables[table]['CONG_COND'][tablerow] = "UNDISTURBED"
                        if "#" in str(self.tables[table]['CONG_PDEN'][tablerow].lower()):
                            self.tables[table]['CONG_PDEN'][tablerow] = str(self.tables[table]['CONG_PDEN'][tablerow]).rsplit('#', 2)[1]

            except Exception as e:
                print(f"Couldn't find table or field, skipping... {str(e)}")
                pass

        self.remove_match_id()
        self.check_matched_to_gint()
        self.enable_buttons()

    
    def match_unique_id_psl(self):
        self.disable_buttons()
        self.get_gint()
        self.matched = False
        self.error = False

        if not self.gint_location or self.gint_location == '':
            self.text.setText('''AGS file loaded.
''')
            QApplication.processEvents()
            return

        self.text.setText('''Matching AGS to gINT, please wait...
''')
        QApplication.processEvents()
        print(f"Matching PSL AGS to gINT... {self.gint_location}") 

        self.get_ags_tables()

        self.get_spec()['Depth'] = self.get_spec()['Depth'].map('{:,.2f}'.format)
        self.get_spec()['Depth'] = self.get_spec()['Depth'].astype(str)
        self.get_spec()['match_id'] = self.get_spec()['PointID']
        self.get_spec()['match_id'] += self.get_spec()['Depth']

        for table in self.ags_tables:
            try:
                gint_rows = self.get_spec().shape[0]

                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']

                try:
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                self.matched = True
                                self.tables[table]['LOCA_ID'][tablerow] = self.get_spec()['PointID'][gintrow]
                                self.tables[table]['SAMP_ID'][tablerow] = self.get_spec()['SAMP_ID'][gintrow]
                                self.tables[table]['SAMP_REF'][tablerow] = self.get_spec()['SAMP_REF'][gintrow]
                                self.tables[table]['SAMP_TYPE'][tablerow] = self.get_spec()['SAMP_TYPE'][gintrow]
                                self.tables[table]['SPEC_REF'][tablerow] = self.get_spec()['SPEC_REF'][gintrow]
                                self.tables[table]['SAMP_TOP'][tablerow] = format(self.get_spec()['SAMP_Depth'][gintrow],'.2f')
                                self.tables[table]['SPEC_DPTH'][tablerow] = self.get_spec()['Depth'][gintrow]
                                
                                for x in self.tables[table].keys():
                                    if "LAB" in x:
                                        self.tables[table][x][tablerow] = "PSL"
                except:
                    pass

                '''CONG'''
                if table == 'CONG':
                    for tablerow in range(2,len(self.tables[table])):
                        if "undisturbed" in str(self.tables[table]['CONG_COND'][tablerow].lower()):
                            self.tables[table]['CONG_COND'][tablerow] = "UNDISTURBED"
                        if "oed" in str(self.tables[table]['CONG_TYPE'][tablerow].lower()):
                            self.tables[table]['CONG_TYPE'][tablerow] = "IL OEDOMETER"
                            self.tables[table]['CONG_COND'][tablerow] = "UNDISTURBED"


                '''TREG'''
                if table == 'TREG':
                    for tablerow in range(2,len(self.tables[table])):
                        if "undisturbed" in str(self.tables[table]['TREG_COND'][tablerow].lower()):
                            self.tables[table]['TREG_COND'][tablerow] = "UNDISTURBED"


                '''TRET'''
                if table == 'TRET':
                    for tablerow in range(2,len(self.tables[table])):
                        if 'TRET_SHST' not in self.tables[table].keys():
                            self.tables[table].insert(len(self.tables[table].keys()),'TRET_SHST','')
                        if self.tables[table]['TRET_SHST'][tablerow] == '' and self.tables[table]['TRET_DEVF'][tablerow] != '':
                            self.tables[table]['TRET_SHST'][tablerow] = round(float(self.tables[table]['TRET_DEVF'][tablerow]) / 2)


                '''PTST'''
                if table == 'PTST':
                    for tablerow in range(2,len(self.tables[table])):
                        if "#" in str(self.tables[table]['PTST_PDEN'][tablerow].lower()):
                            self.tables[table]['PTST_PDEN'][tablerow] = str(self.tables[table]['PTST_PDEN'][tablerow]).rsplit('#', 2)[1]
                        if "undisturbed" in str(self.tables[table]['PTST_COND'][tablerow].lower()):
                            self.tables[table]['PTST_COND'][tablerow] = "UNDISTURBED"
                        if "remoulded" in str(self.tables[table]['PTST_COND'][tablerow].lower()):
                            self.tables[table]['PTST_COND'][tablerow] = "REMOULDED"
                
            except Exception as e:
                print(f"Couldn't find table or field, skipping... {str(e)}")
                pass

        self.remove_match_id()
        self.check_matched_to_gint()
        self.enable_buttons()


    def match_unique_id_geolabs(self):
        self.disable_buttons()
        self.get_gint()
        self.matched = False
        self.error = False

        if not self.gint_location or self.gint_location == '':
            self.text.setText('''AGS file loaded.
''')
            QApplication.processEvents()
            return

        self.text.setText('''Matching AGS to gINT, please wait...
''')
        QApplication.processEvents()
        print(f"Matching Geolabs AGS to gINT... {self.gint_location}") 

        self.get_ags_tables()

        self.get_spec()['Depth'] = self.get_spec()['Depth'].map('{:,.2f}'.format)
        self.get_spec()['Depth'] = self.get_spec()['Depth'].astype(str)
        self.get_spec()['match_id'] = self.get_spec()['PointID']
        self.get_spec()['match_id'] += self.get_spec()['Depth']

        for table in self.ags_tables:
            try:
                gint_rows = self.get_spec().shape[0]

                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']

                try:
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                self.matched = True
                                self.tables[table]['LOCA_ID'][tablerow] = self.get_spec()['PointID'][gintrow]
                                self.tables[table]['SAMP_ID'][tablerow] = self.get_spec()['SAMP_ID'][gintrow]
                                self.tables[table]['SAMP_REF'][tablerow] = self.get_spec()['SAMP_REF'][gintrow]
                                self.tables[table]['SAMP_TYPE'][tablerow] = self.get_spec()['SAMP_TYPE'][gintrow]
                                self.tables[table]['SPEC_REF'][tablerow] = self.get_spec()['SPEC_REF'][gintrow]
                                self.tables[table]['SAMP_TOP'][tablerow] = format(self.get_spec()['SAMP_Depth'][gintrow],'.2f')
                                self.tables[table]['SPEC_DPTH'][tablerow] = self.get_spec()['Depth'][gintrow]
                                
                                for x in self.tables[table].keys():
                                    if "LAB" in x:
                                        self.tables[table][x][tablerow] = "Geolabs Limited"
                except:
                    pass


                '''PTST'''
                if table == 'PTST':
                    for tablerow in range(2,len(self.tables[table])):
                        if "#" in str(self.tables[table]['PTST_PDEN'][tablerow].lower()):
                            self.tables[table]['PTST_PDEN'][tablerow] = str(self.tables[table]['PTST_PDEN'][tablerow]).rsplit('#', 2)[1]
                        if "undisturbed" in str(self.tables[table]['PTST_COND'][tablerow].lower()):
                            self.tables[table]['PTST_COND'][tablerow] = "UNDISTURBED"
                        if "remoulded" in str(self.tables[table]['PTST_COND'][tablerow].lower()):
                            self.tables[table]['PTST_COND'][tablerow] = "REMOULDED"
                        if str(self.tables[table]['PTST_TESN'][tablerow]) == '':
                            self.tables[table]['PTST_TESN'][tablerow] = "1"

            except Exception as e:
                print(f"Couldn't find table or field, skipping... {str(e)}")
                pass

        self.remove_match_id()
        self.check_matched_to_gint()
        self.enable_buttons()


    def match_unique_id_geolabs_fugro(self):
        self.disable_buttons()
        self.get_gint()
        self.matched = False
        self.error = False

        if not self.gint_location or self.gint_location == '':
            self.text.setText('''AGS file loaded.
''')
            QApplication.processEvents()
            return

        self.text.setText('''Matching AGS to gINT, please wait...
''')
        QApplication.processEvents()
        print(f"Matching Geolabs AGS to gINT... {self.gint_location}") 

        self.get_ags_tables()
        
        '''Using for Fugro Boreholes (50HZ samples have different SAMP format including dupe depths)'''
        self.get_spec()['SAMP_Depth'] = self.get_spec()['SAMP_Depth'].map('{:,.2f}'.format)
        self.get_spec()['SAMP_Depth'] = self.get_spec()['SAMP_Depth'].astype(str)
        self.get_spec()['match_id'] = self.get_spec()['PointID']
        self.get_spec()['match_id'] += self.get_spec()['SAMP_Depth']
        self.get_spec()['match_id'] += self.get_spec()['SAMP_REF']

        for table in self.ags_tables:
            try:                
                if 'Depth' not in self.tables[table]:
                    self.tables[table].insert(8,'Depth','')

                gint_rows = self.get_spec().shape[0]

                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_REF']

                try:
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                self.matched = True
                                self.tables[table]['Depth'][tablerow] = self.tables[table]['SPEC_DPTH'][tablerow]
                                self.tables[table]['LOCA_ID'][tablerow] = self.get_spec()['PointID'][gintrow]
                                self.tables[table]['SAMP_ID'][tablerow] = self.get_spec()['SAMP_ID'][gintrow]
                                self.tables[table]['SAMP_REF'][tablerow] = self.get_spec()['SAMP_REF'][gintrow]
                                self.tables[table]['SAMP_TYPE'][tablerow] = self.get_spec()['SAMP_TYPE'][gintrow]
                                self.tables[table]['SPEC_REF'][tablerow] = self.get_spec()['SPEC_REF'][gintrow]
                                self.tables[table]['SAMP_TOP'][tablerow] = self.get_spec()['SAMP_Depth'][gintrow]
                                self.tables[table]['SPEC_DPTH'][tablerow] = format(self.get_spec()['Depth'][gintrow],'.2f')
                                
                                # for x in self.tables[table].keys():
                                #     if "LAB" in x:
                                #         self.tables[table][x][tablerow] = "Geolabs"

                except Exception as e:
                    print(e)
                    pass

                '''RPLT'''
                if table == 'RPLT':
                    for tablerow in range(2,len(self.tables[table])):
                        if self.tables[table]['match_id'][tablerow] == self.tables[table]['match_id'][tablerow -1]:
                            self.tables[table]['Depth'][tablerow] = format(float(self.tables[table]['Depth'][tablerow]) + 0.01,'.2f')
                        if self.tables[table]['match_id'][tablerow] == self.tables[table]['match_id'][tablerow -2]:
                            self.tables[table]['Depth'][tablerow] = format(float(self.tables[table]['Depth'][tablerow]) + 0.01,'.2f')
                        try:
                            if self.tables[table]['match_id'][tablerow] == self.tables[table]['match_id'][tablerow -3]:
                                self.tables[table]['Depth'][tablerow] = format(float(self.tables[table]['Depth'][tablerow]) + 0.01,'.2f')
                        except:
                            pass


                '''PTST'''
                if table == 'PTST':
                    for tablerow in range(2,len(self.tables[table])):
                        if "#" in str(self.tables[table]['PTST_PDEN'][tablerow].lower()):
                            self.tables[table]['PTST_PDEN'][tablerow] = str(self.tables[table]['PTST_PDEN'][tablerow]).rsplit('#', 2)[1]
                        if "undisturbed" in str(self.tables[table]['PTST_COND'][tablerow].lower()):
                            self.tables[table]['PTST_COND'][tablerow] = "UNDISTURBED"
                        if "remoulded" in str(self.tables[table]['PTST_COND'][tablerow].lower()):
                            self.tables[table]['PTST_COND'][tablerow] = "REMOULDED"
                        if str(self.tables[table]['PTST_TESN'][tablerow]) == '':
                            self.tables[table]['PTST_TESN'][tablerow] = "1"                

            except Exception as e:
                print(f"Couldn't find table or field, skipping... {str(e)}")
                pass

        self.remove_match_id()
        self.check_matched_to_gint()
        self.enable_buttons()


    def match_unique_id_gqm_pez(self):
        self.disable_buttons()
        self.get_gint()
        self.matched = False
        self.error = False

        if not self.gint_location or self.gint_location == '':
            self.text.setText('''AGS file loaded.
''')
            QApplication.processEvents()
            return

        self.text.setText('''Matching AGS to gINT, please wait...
''')
        QApplication.processEvents()
        print(f"Matching GM Lab AGS to gINT... {self.gint_location}") 

        self.get_ags_tables()

        if 'GCHM' in self.ags_tables or 'ERES' in self.ags_tables:
            self.error = True
            print("GCHM or ERES table(s) found.")

        self.get_spec()['Depth'] = self.get_spec()['Depth'].map('{:,.2f}'.format)
        self.get_spec()['Depth'] = self.get_spec()['Depth'].astype(str)
        self.get_spec()['match_id'] = self.get_spec()['PointID']
        self.get_spec()['match_id'] += self.get_spec()['SPEC_REF']
        self.get_spec()['match_id'] += self.get_spec()['Depth']
        self.get_spec()['batched'] = self.get_spec()['SAMP_TYPE'].astype(str).str[0]
        self.get_spec()['match_id'] += self.get_spec()['batched']
        self.get_spec().drop(['batched'], axis=1, inplace=True)

        for table in self.ags_tables:
            try:
                gint_rows = self.get_spec().shape[0]

                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TYPE']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']
                self.tables[table]['batched'] = self.tables[table]['SAMP_REF'].astype(str).str[0]
                self.tables[table]['match_id'] += self.tables[table]['batched']
                self.tables[table].drop(['batched'], axis=1, inplace=True)
                    
                try:
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):

                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                self.matched = True

                                if table == 'CONG':
                                    if self.tables[table]['SPEC_REF'][tablerow] == "OED" or self.tables[table]['SPEC_REF'][tablerow] == "OEDR" and self.tables[table]['CONG_TYPE'][tablerow] == '':
                                        self.tables[table]['CONG_TYPE'][tablerow] = self.tables[table]['SPEC_REF'][tablerow]

                                if table == 'SAMP':
                                    self.tables[table]['SAMP_REM'][tablerow] = self.get_spec()['SPEC_REF'][gintrow]

                                self.tables[table]['SAMP_ID'][tablerow] = self.get_spec()['SAMP_ID'][gintrow]
                                self.tables[table]['SAMP_REF'][tablerow] = self.get_spec()['SAMP_REF'][gintrow]
                                self.tables[table]['SAMP_TYPE'][tablerow] = self.get_spec()['SAMP_TYPE'][gintrow]
                                self.tables[table]['SAMP_TOP'][tablerow] = format(self.get_spec()['SAMP_Depth'][gintrow],'.2f')

                                try:
                                    self.tables[table]['SPEC_REF'][tablerow] = self.get_spec()['SPEC_REF'][gintrow]
                                except:
                                    pass

                                try:
                                    self.tables[table]['SPEC_DPTH'][tablerow] = self.get_spec()['Depth'][gintrow]
                                except:
                                    pass

                                for x in self.tables[table].keys():
                                    if "LAB" in x:
                                        self.tables[table][x][tablerow] = "GM Lab"

                except Exception as e:
                    print(str(e))
                    pass

                '''SHBG'''
                if table == 'SHBG':
                    for tablerow in range(2,len(self.tables[table])):
                        if "small" in str(self.tables[table]['SHBG_TYPE'][tablerow].lower()):
                            self.tables[table]['SHBG_REM'][tablerow] += " - " + self.tables[table]['SHBG_TYPE'][tablerow]
                            self.tables[table]['SHBG_TYPE'][tablerow] = "SMALL SBOX"

                
                '''SHBT'''
                if table == 'SHBT':
                    for tablerow in range(2,len(self.tables[table])):
                        if self.tables[table]['SHBT_NORM'][tablerow]:
                            self.tables[table]['SHBT_NORM'][tablerow] = round(float(self.tables[table]['SHBT_NORM'][tablerow]))


                '''LLPL'''
                if table == 'LLPL':
                    if 'Non-Plastic' not in self.tables[table]:
                        self.tables[table].insert(13,'Non-Plastic','')
                    for tablerow in range(2,len(self.tables[table])):
                        if self.tables[table]['LLPL_LL'][tablerow] == '' and self.tables[table]['LLPL_PL'][tablerow] == '' and self.tables[table]['LLPL_PI'][tablerow] == '':
                            self.tables[table]['Non-Plastic'][tablerow] = -1


                '''GRAG'''
                if table == 'GRAG':
                    for tablerow in range(2,len(self.tables[table])):
                        if self.tables['GRAG']['GRAG_SILT'][tablerow] == '' and self.tables['GRAG']['GRAG_CLAY'][tablerow] == '':
                            if self.tables['GRAG']['GRAG_VCRE'][tablerow] == '':
                                self.tables['GRAG']['GRAG_FINE'][tablerow] = format(100 - (float(self.tables['GRAG']['GRAG_GRAV'][tablerow])) - (float(self.tables['GRAG']['GRAG_SAND'][tablerow])),".1f")
                            else:
                                self.tables['GRAG']['GRAG_FINE'][tablerow] = format(100 - (float(self.tables['GRAG']['GRAG_VCRE'][tablerow])) - (float(self.tables['GRAG']['GRAG_GRAV'][tablerow])) - (float(self.tables['GRAG']['GRAG_SAND'][tablerow])),'.1f')
                        else:
                            self.tables['GRAG']['GRAG_FINE'][tablerow] = format((float(self.tables['GRAG']['GRAG_SILT'][tablerow]) + float(self.tables['GRAG']['GRAG_CLAY'][tablerow])),'.1f')


                '''GRAT'''
                if table == 'GRAT':
                    for tablerow in range(2,len(self.tables[table])):
                        if self.tables[table]['GRAT_PERP'][tablerow]:
                            self.tables[table]['GRAT_PERP'][tablerow] = round(float(self.tables[table]['GRAT_PERP'][tablerow]))


                '''TREG'''
                if table == 'TREG':
                    for tablerow in range(2,len(self.tables[table])):
                        if self.tables[table]['TREG_TYPE'][tablerow] == 'CU' and self.tables[table]['TREG_COH'][tablerow] == '0':
                            self.tables[table]['TREG_COH'][tablerow] = ''
                            self.tables[table]['TREG_PHI'][tablerow] = ''
                            self.tables[table]['TREG_COND'][tablerow] = 'UNDISTURBED'
                        if self.tables[table]['TREG_TYPE'][tablerow] == 'CD':
                            self.tables[table]['TREG_COND'][tablerow] = 'REMOULDED'
                            if self.tables[table]['TREG_PHI'][tablerow] == '':
                                cid_sample = str(self.tables[table]['SAMP_ID'][tablerow]) + "-" + str(self.tables[table]['SPEC_REF'][tablerow])
                                print(f'CID result: {cid_sample} - does not have friction angle.')


                '''TRET'''
                if table == 'TRET':
                    for tablerow in range(2,len(self.tables[table])):
                        if 'TRET_SHST' in self.tables[table].keys():
                            if self.tables[table]['TRET_SHST'][tablerow] == '' and self.tables[table]['TRET_DEVF'][tablerow] != '':
                                if "cell" in str(self.tables['TRET']['TRET_SAT'][tablerow]).lower():
                                    self.tables[table]['TRET_SHST'][tablerow] = round(float(self.tables[table]['TRET_DEVF'][tablerow]) / 2)
                        if 'TRET_CELL' in self.tables[table].keys():
                            if not self.tables[table]['TRET_CELL'][tablerow] == '':
                                self.tables[table]['TRET_CELL'][tablerow] = round(float(self.tables[table]['TRET_CELL'][tablerow]))

                '''LPDN'''
                if table == 'LPDN':
                    for tablerow in range(2,len(self.tables[table])):
                        if self.tables[table]['LPDN_TYPE'][tablerow] == 'LARGE PKY':
                            self.tables[table]['LPDN_TYPE'][tablerow] = 'LARGE PYK'


                '''CONG'''
                if table == 'CONG':
                    for tablerow in range(2,len(self.tables[table])):
                        if self.tables[table]['CONG_TYPE'][tablerow] == '' and self.tables[table]['CONG_COND'][tablerow] == 'Intact':
                            self.tables[table]['CONG_TYPE'][tablerow] = 'CRS'
                            self.tables[table]['CONG_COND'][tablerow] = 'UNDISTURBED'
                        if "intact" in str(self.tables[table]['CONG_COND'][tablerow].lower()):
                            self.tables[table]['CONG_COND'][tablerow] = "UNDISTURBED"
                        if "oed" in str(self.tables[table]['CONG_TYPE'][tablerow].lower()):
                            self.tables[table]['CONG_TYPE'][tablerow] = "IL OEDOMETER"
                            self.tables[table]['CONG_COND'][tablerow] = "UNDISTURBED"
                        self.tables[table]['CONG_COND'][tablerow] = str(self.tables[table]['CONG_COND'][tablerow].upper())


                '''TRIG & TRIT'''
                if table == 'TRIG' or table == 'TRIT':
                    if 'Depth' not in self.tables[table]:
                        self.tables[table].insert(8,'Depth','')
                    if table == 'TRIT':
                        for tablerow in range(2,len(self.tables[table])):
                            if self.tables[table]['TRIT_DEVF'][tablerow]:
                                self.tables[table]['TRIT_DEVF'][tablerow] = round(float(self.tables[table]['TRIT_DEVF'][tablerow]))
                            if self.tables[table]['TRIT_TESN'][tablerow] == '':
                                self.tables[table]['TRIT_TESN'][tablerow] = 1
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                if self.tables['TRIG']['TRIG_COND'][tablerow] == 'REMOULDED':
                                    self.tables[table]['Depth'][tablerow] = float(self.get_spec()['Depth'][gintrow]) + 0.01
                                else:
                                    self.tables[table]['Depth'][tablerow] = self.get_spec()['Depth'][gintrow]


                '''RELD'''
                if table == 'RELD':
                    if 'Depth' not in self.tables[table]:
                        self.tables[table].insert(8,'Depth','')
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                self.tables[table]['Depth'][tablerow] = self.get_spec()['Depth'][gintrow]


                '''LDYN'''
                if table == 'LDYN':
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                if 'LDYN_SWAV1' in self.tables[table] or 'LDYN_SWAV1SS' in self.tables[table]:
                                    if self.tables[table]['LDYN_SWAV1SS'][tablerow] == "":
                                        if self.tables[table]['LDYN_SWAV5'][tablerow] == "":
                                            self.tables[table]['LDYN_SWAV'][tablerow] = int(mean([int(float(self.tables[table]['LDYN_SWAV1'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV2'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV3'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV4'][tablerow]))
                                            ]))
                                        else:
                                            self.tables[table]['LDYN_SWAV'][tablerow] = int(mean([int(float(self.tables[table]['LDYN_SWAV1'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV2'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV3'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV4'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV5'][tablerow]))
                                            ]))
                                    else:
                                        if self.tables[table]['LDYN_SWAV5SS'][tablerow] == "":
                                            self.tables[table]['LDYN_SWAV'][tablerow] = int(mean([int(float(self.tables[table]['LDYN_SWAV1SS'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV2SS'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV3SS'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV4SS'][tablerow]))
                                            ]))
                                        else:
                                            self.tables[table]['LDYN_SWAV'][tablerow] = int(mean([int(float(self.tables[table]['LDYN_SWAV1SS'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV2SS'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV3SS'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV4SS'][tablerow])),
                                            int(float(self.tables[table]['LDYN_SWAV5SS'][tablerow]))
                                            ]))
                            if self.tables[table]['LDYN_REM'][tablerow] == "":
                                self.tables[table]['LDYN_REM'][tablerow] = "Bender Element"

            except Exception as e:
                print(f"Couldn't find table or field, skipping... {str(e)}")
                pass

        self.remove_match_id()
        self.check_matched_to_gint()
        self.enable_buttons()


    def match_unique_id_dets_pez(self):
        self.disable_buttons()
        self.get_gint()
        self.matched = False
        self.error = False

        if not self.gint_location or self.gint_location == '':
            self.text.setText('''AGS file loaded.
''')
            QApplication.processEvents()
            return

        self.text.setText('''Matching AGS to gINT, please wait...
''')
        QApplication.processEvents()
        print(f"Matching DETS for PEZ AGS to gINT... {self.gint_location}") 

        self.get_ags_tables()

        if 'GCHM' in self.ags_tables or 'ERES' in self.ags_tables:
            pass
        else:
            self.error = True
            print("Cannot find GCHM or ERES - looks like this AGS is from GM Lab.")

        self.create_match_id()

        for table in self.ags_tables:
            try:
                gint_rows = self.get_spec().shape[0]

                for row in range (0,gint_rows):
                    self.get_spec()['match_id'][row] = str(self.get_spec()['PointID'][row]) + str(format(self.get_spec()['Depth'][row],'.2f')) + str(self.get_spec()['SAMP_TYPE'][row][0]) + str(self.get_spec()['SPEC_REF'][row])

                for row in range (2,len(self.tables[table])):
                    self.tables[table]['match_id'][row] = str(self.tables[table]['LOCA_ID'][row]).rsplit(' ', 2)[0] + str(self.tables[table]['SAMP_TOP'][row]) + str(self.tables[table]['SAMP_REF'][row][0]) + str(self.tables[table]['SAMP_REF'][row]).split(' ')[1]

                try:
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                self.matched = True
                                if table == 'ERES':
                                    if 'ERES_REM' not in self.tables[table].keys():
                                        self.tables[table].insert(len(self.tables[table].keys()),'ERES_REM','')
                                    self.tables[table]['ERES_REM'][tablerow] = self.tables[table]['SPEC_REF'][tablerow]
                                self.tables[table]['LOCA_ID'][tablerow] = self.get_spec()['PointID'][gintrow]
                                self.tables[table]['SAMP_ID'][tablerow] = self.get_spec()['SAMP_ID'][gintrow]
                                self.tables[table]['SAMP_REF'][tablerow] = self.get_spec()['SAMP_REF'][gintrow]
                                self.tables[table]['SAMP_TYPE'][tablerow] = self.get_spec()['SAMP_TYPE'][gintrow]
                                self.tables[table]['SPEC_REF'][tablerow] = self.get_spec()['SPEC_REF'][gintrow]
                                self.tables[table]['SAMP_TOP'][tablerow] = format(self.get_spec()['SAMP_Depth'][gintrow],'.2f')
                                self.tables[table]['SPEC_DPTH'][tablerow] = format(self.get_spec()['Depth'][gintrow],'.2f')
                                
                                for x in self.tables[table].keys():
                                    if "LAB" in x:
                                        self.tables[table][x][tablerow] = "DETS"
                except Exception as e:
                    print(e)
                    pass

                '''GCHM'''
                if table == 'GCHM':
                    for tablerow in range(2,len(self.tables[table])):
                        if "ph" in str(self.tables[table]['GCHM_UNIT'][tablerow].lower()):
                            self.tables[table]['GCHM_UNIT'][tablerow] = "-"
                        if "co3" in str(self.tables[table]['GCHM_CODE'][tablerow].lower()):
                            self.tables[table]['GCHM_CODE'][tablerow] = "CACO3"


                '''ERES'''
                if table == 'ERES':
                    for tablerow in range(2,len(self.tables[table])):
                        if "<" in str(self.tables[table]['ERES_RTXT'][tablerow].lower()):
                            self.tables[table]['ERES_RTXT'][tablerow] = str(self.tables[table]['ERES_RTXT'][tablerow]).rsplit(" ", 1)[1]
                        if "solid_21" in str(self.tables[table]['ERES_REM'][tablerow].lower()) or "2:1" in str(self.tables[table]['ERES_NAME'][tablerow].lower()):
                            self.tables[table]['ERES_NAME'][tablerow] = "SOLID_21 WATER EXTRACT"
                        if "solid_wat" in str(self.tables[table]['ERES_REM'][tablerow].lower()):
                            self.tables[table]['ERES_NAME'][tablerow] = "SOLID_11 WATER EXTRACT"
                        if "solid_tot" in str(self.tables[table]['ERES_REM'][tablerow].lower()):
                            self.tables[table]['ERES_NAME'][tablerow] = "SOLID_TOTAL"
                        if "sulph" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()) and "so4" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()) or "sulf" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                            self.tables[table]['ERES_TNAM'][tablerow] = "WS"
                        if "sulph" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()) and "total" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                            self.tables[table]['ERES_TNAM'][tablerow] = "TS"
                        if "caco3" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                            self.tables[table]['ERES_TNAM'][tablerow] = "CACO3"
                        if "co2" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                            self.tables[table]['ERES_TNAM'][tablerow] = "CO2"
                        if "ph" == str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                            self.tables[table]['ERES_TNAM'][tablerow] = "PH"
                        if "chloride" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                            self.tables[table]['ERES_TNAM'][tablerow] = "Cl"
                        if "los" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                            self.tables[table]['ERES_TNAM'][tablerow] = "LOI"
                        if "ph" in str(self.tables[table]['ERES_RUNI'][tablerow].lower()):
                            self.tables[table]['ERES_RUNI'][tablerow] = "-"

            except Exception as e:
                print(f"Couldn't find table or field, skipping... {str(e)}")
                pass

        self.remove_match_id()
        self.check_matched_to_gint()
        self.enable_buttons()



    def ags_table_reset(self):
        if not self.ags_tables == []:
            self.ags_tables = []
        return self.ags_tables

        
    def del_non_lab_tables(self):
        self.get_ags_tables()

        for table in self.result_tables:
            if table in list(self.tables):
                self.ags_tables.append(table)

        for table in list(self.tables):
            if table not in self.ags_tables and not table == 'TRAN' and not table == 'PROJ':
                del self.tables[table]
                print(f"{str(table)} table deleted.")
            try:
                if table == 'SAMP' or table == 'SPEC':
                    del self.tables[table]
                    print(f"{str(table)} table deleted.")
            except:
                pass


    def get_cpt_tables(self):
        self.ags_table_reset()

        self.cpt_tables = ["SCPG","SCPT","SCPP","SCCG","SCCT","SCDG","SCDT"]

        for table in self.cpt_tables:
            if table in list(self.tables):
                self.ags_tables.append(table)


    def del_non_cpt_tables(self):
        self.get_cpt_tables()

        if not self.ags_tables == []:
            for table in self.core_tables:
                if table in list(self.tables):
                    self.ags_tables.append(table)

            for table in list(self.tables):
                if table not in self.ags_tables:
                    del self.tables[table]
                    print(f"{str(table)} table deleted.")

            self.text.setText('''CPT Only export ready.
Click "Save AGS file"''')
            QApplication.processEvents()
            print("CPT Data export ready. Click 'Save AGS file'.")

        else:
            self.text.setText('''Could not find any CPT tables.
Check the AGS with "View data".''')
            QApplication.processEvents()
            print("No CPT groups found - did this AGS contain CPT data? Check the data with 'View data'.")

    
    def get_lab_tables(self):
        self.get_ags_tables()
        self.result_tables.append('GEOL')
        self.result_tables.append('DREM')
        self.result_tables.append('DETL')

        for table in self.result_tables:
            if table in list(self.tables):
                self.ags_tables.append(table)

        for table in list(self.result_tables):
            if table == 'GEOL' or table == 'DREM' or table == 'DETL':
                self.result_tables.remove(table)


    def export_lab_only(self):
        self.get_lab_tables()

        if not self.ags_tables == []:

            for table in self.core_tables:
                if table in list(self.tables):
                    self.ags_tables.append(table)

            for table in list(self.tables):
                if table not in self.ags_tables:
                    del self.tables[table]
                    print(f"{str(table)} table deleted.")

            self.text.setText('''Lab Data & GEOL export ready.
Click "Save AGS file"''')
            QApplication.processEvents()
            print("Lab Data & GEOL export ready. Click 'Save AGS file'.")
            self.ags_table_reset()

        else:
            self.text.setText('''Could not find any Lab or GEOL tables.
Check the AGS with "View data".''')
            QApplication.processEvents()
            print("No Lab or GEOL groups found - did this AGS contain CPT data? Check the data with 'View data'.")
            self.ags_table_reset()

      
    # def play_coin(self):
    #     coin_num = np.random.randint(77)
    #     if coin_num == 17:
    #         coin = ('assets/sounds/coin.mp3')
    #         coin_url = QUrl.fromLocalFile(coin)
    #         content = QMediaContent(coin_url)
    #         self.player.setMedia(content)
    #         self.player.setVolume(33)
    #         self.player.play()

    # def play_nice(self):
    #     nice_num = np.random.randint(77)
    #     if nice_num == 7:
    #         nice = ('assets/sounds/nice.mp3')
    #         nice_url = QUrl.fromLocalFile(nice)
    #         content = QMediaContent(nice_url)
    #         self.player.setMedia(content)
    #         self.player.setVolume(33)
    #         self.player.play()

    # def dark_toggle(self):
    #     self.play_nice()
    #     self.dark_mode()

    def dark_mode(self):
        if self.dark_mode_button.isChecked() == False:
            #LIGHT THEME
            self.dark_mode_button.setChecked(False)
            self.config.set('Theme','dark','')
            with open('common/assets/settings.ini', 'w') as configfile: 
                self.config.write(configfile)
            self.plot_area.setBackground("#f0f0f0")
            #self.button_copy_actual.setIcon(QtGui.QIcon('assets/images/copy.png'))
            #self.button_copy_avg.setIcon(QtGui.QIcon('assets/images/copy.png'))
            
            light_palette = QPalette()
            light_palette.setColor(QPalette.Window, QColor(240, 240, 240))
            light_palette.setColor(QPalette.WindowText, Qt.black)
            light_palette.setColor(QPalette.Base, QColor(240, 240, 240))
            light_palette.setColor(QPalette.AlternateBase, QColor(240, 240, 240))
            light_palette.setColor(QPalette.ToolTipBase, QColor(240, 240, 240))
            light_palette.setColor(QPalette.ToolTipText, Qt.black)
            light_palette.setColor(QPalette.Text, Qt.black)#
            light_palette.setColor(QPalette.Button, QColor(240, 240, 240))
            light_palette.setColor(QPalette.ButtonText, Qt.white)#
            light_palette.setColor(QPalette.BrightText, Qt.red)
            light_palette.setColor(QPalette.Link, QColor(42, 130, 218))
            light_palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
            light_palette.setColor(QPalette.HighlightedText, QColor(240, 240, 240))
            light_palette.setColor(QPalette.Active, QPalette.Button, QColor(240, 240, 240))
            light_palette.setColor(QPalette.Disabled, QPalette.ButtonText, Qt.lightGray)
            light_palette.setColor(QPalette.Disabled, QPalette.WindowText, Qt.lightGray)
            light_palette.setColor(QPalette.Disabled, QPalette.Text, Qt.lightGray)
            light_palette.setColor(QPalette.Disabled, QPalette.Light, QColor('#f0f0f0'))
            self.left_spacer.setStyleSheet(f"{self.config.get('Theme','spacer_css_light')}")
            self.top_spacer.setStyleSheet(f"{self.config.get('Theme','spacer_css_light')}")
            self.top_spacer.setStyleSheet(f"{self.config.get('Theme','spacer_css_light')}")
            self.right_spacer.setStyleSheet(f"{self.config.get('Theme','spacer_css_light')}")
            self.bot_spacer.setStyleSheet(f"{self.config.get('Theme','spacer_css_light')}")
            self.unit_textbox.setStyleSheet(f"{self.config.get('Theme','textbox_css_light')}")
            self.actual_val.setStyleSheet(f"{self.config.get('Theme','textbox_dark_css_light')}")
            self.average_val.setStyleSheet(f"{self.config.get('Theme','textbox_dark_css_light')}")
            self.button_copy_actual.setStyleSheet(f"{self.config.get('Theme','button_transp_css_light')}")
            self.button_copy_avg.setStyleSheet(f"{self.config.get('Theme','button_transp_css_light')}")
            self.button_gint.setStyleSheet(f"{self.config.get('Theme','button_css_light')}")
            self.button_depth.setStyleSheet(f"{self.config.get('Theme','button_css_light')}")
            self.button_cpt_val.setStyleSheet(f"{self.config.get('Theme','button_css_light')}")
            self.remove_before.setStyleSheet(f"{self.config.get('Theme','button_css_sml_light')}")
            self.remove_after.setStyleSheet(f"{self.config.get('Theme','button_css_sml_light')}")
            self.remove_at.setStyleSheet(f"{self.config.get('Theme','button_css_sml_light')}")
            self.re_plot.setStyleSheet(f"{self.config.get('Theme','button_css_sml_light')}")
            self.increment.setStyleSheet(f"{self.config.get('Theme','button_css_sml_light')}")
            self.decrement.setStyleSheet(f"{self.config.get('Theme','button_css_sml_light')}")
            self.cpt_table.setStyleSheet(f"{self.config.get('Theme','combo_css_light')}")
            self.geol_layers.setStyleSheet(f"{self.config.get('Theme','combo_css_light')}")
            self.avg_vals.setStyleSheet(f"{self.config.get('Theme','combo_css_light')}")
            self.dark_mode_button.setStyleSheet(f"{self.config.get('Theme','checkbox_css_light')}")
            self.full_bh.setStyleSheet(f"{self.config.get('Theme','checkbox_css_light')}")
            self.menubar.setStyleSheet(f"font: 10pt 'Roboto'; background: #f0f0f0; color: black;")
            self.point_table.setStyleSheet(f"{self.config.get('Theme','table_css_light')}")
            self.depth_table.setStyleSheet(f"{self.config.get('Theme','table_css_light')}")

            self.reset_graph()   
            if not self.cpt_value == "":
                self.plot_graph(x=self.x, y=self.y, cpt_value=self.cpt_value)
            QApplication.setPalette(light_palette)


        else:
            #DARK THEME
            self.dark_mode_button.setChecked(True)
            self.config.set('Theme','dark','True')
            with open('common/assets/settings.ini', 'w') as configfile: 
                self.config.write(configfile)
            self.plot_area.setBackground("#353535")
            #self.button_copy_actual.setIcon(QtGui.QIcon('assets/images/copy_light.png'))
            #self.button_copy_avg.setIcon(QtGui.QIcon('assets/images/copy_light.png'))

            dark_palette = QPalette()
            dark_palette.setColor(QPalette.Window, QColor(53, 53, 53))
            dark_palette.setColor(QPalette.WindowText, Qt.black)
            dark_palette.setColor(QPalette.Base, QColor(35, 35, 35))
            dark_palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
            dark_palette.setColor(QPalette.ToolTipBase, QColor(25, 25, 25))
            dark_palette.setColor(QPalette.ToolTipText, Qt.black)
            dark_palette.setColor(QPalette.Text, Qt.white)
            dark_palette.setColor(QPalette.Button, QColor(53, 53, 53))
            dark_palette.setColor(QPalette.ButtonText, Qt.black)
            dark_palette.setColor(QPalette.BrightText, Qt.red)
            dark_palette.setColor(QPalette.Link, QColor(42, 130, 218))
            dark_palette.setColor(QPalette.Highlight, Qt.darkGray)
            dark_palette.setColor(QPalette.HighlightedText, QColor(35, 35, 35))
            dark_palette.setColor(QPalette.Active, QPalette.Button, QColor(53, 53, 53))
            dark_palette.setColor(QPalette.Disabled, QPalette.ButtonText, Qt.darkGray)
            dark_palette.setColor(QPalette.Disabled, QPalette.WindowText, Qt.darkGray)
            dark_palette.setColor(QPalette.Disabled, QPalette.Text, Qt.darkGray)
            dark_palette.setColor(QPalette.Disabled, QPalette.Light, QColor(53, 53, 53))
            self.left_spacer.setStyleSheet(f"{self.config.get('Theme','spacer_css')}")
            self.top_spacer.setStyleSheet(f"{self.config.get('Theme','spacer_css')}")
            self.top_spacer.setStyleSheet(f"{self.config.get('Theme','spacer_css')}")
            self.right_spacer.setStyleSheet(f"{self.config.get('Theme','spacer_css')}")
            self.bot_spacer.setStyleSheet(f"{self.config.get('Theme','spacer_css')}")
            self.unit_textbox.setStyleSheet(f"{self.config.get('Theme','textbox_css')}")
            self.actual_val.setStyleSheet(f"{self.config.get('Theme','textbox_dark_css')}")
            self.average_val.setStyleSheet(f"{self.config.get('Theme','textbox_dark_css')}")
            self.button_copy_actual.setStyleSheet(f"{self.config.get('Theme','button_transp_css')}")
            self.button_copy_avg.setStyleSheet(f"{self.config.get('Theme','button_transp_css')}")
            self.button_gint.setStyleSheet(f"{self.config.get('Theme','button_css')}")
            self.button_depth.setStyleSheet(f"{self.config.get('Theme','button_css')}")
            self.button_cpt_val.setStyleSheet(f"{self.config.get('Theme','button_css')}")
            self.remove_before.setStyleSheet(f"{self.config.get('Theme','button_css_sml')}")
            self.remove_after.setStyleSheet(f"{self.config.get('Theme','button_css_sml')}")
            self.remove_at.setStyleSheet(f"{self.config.get('Theme','button_css_sml')}")
            self.re_plot.setStyleSheet(f"{self.config.get('Theme','button_css_sml')}")
            self.increment.setStyleSheet(f"{self.config.get('Theme','button_css_sml')}")
            self.decrement.setStyleSheet(f"{self.config.get('Theme','button_css_sml')}")
            self.cpt_table.setStyleSheet(f"{self.config.get('Theme','combo_css')}")
            self.geol_layers.setStyleSheet(f"{self.config.get('Theme','combo_css')}")
            self.avg_vals.setStyleSheet(f"{self.config.get('Theme','combo_css')}")
            self.dark_mode_button.setStyleSheet(f"{self.config.get('Theme','checkbox_css')}")
            self.full_bh.setStyleSheet(f"{self.config.get('Theme','checkbox_css')}")
            self.menubar.setStyleSheet(f"font: 10pt 'Roboto'; background: #353535; color: white;")
            self.point_table.setStyleSheet(f"{self.config.get('Theme','table_css')}")
            self.depth_table.setStyleSheet(f"{self.config.get('Theme','table_css')}")

            self.reset_graph()
            if not self.cpt_value == "":
                self.plot_graph(x=self.x, y=self.y, cpt_value=self.cpt_value)
            QApplication.setPalette(dark_palette)


    def disable_buttons(self):       
        self.button_open.setEnabled(False)
        self.pandas_gui.setEnabled(False)
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


    def enable_buttons(self):
        self.button_open.setEnabled(True)
        self.pandas_gui.setEnabled(True)
        self.button_count_results.setEnabled(True)
        self.button_ags_checker.setEnabled(True)
        self.button_save_ags.setEnabled(True)
        self.button_del_tbl.setEnabled(True)
        self.button_cpt_only.setEnabled(True)
        self.button_lab_only.setEnabled(True)
        self.lab_select.setEnabled(True)
        self.button_match_lab.setEnabled(True)


    def eventFilter(self, object: QObject, event: QEvent) -> bool:
        if event.type() == QEvent.Type.WindowStateChange:
            maximized = bool(Qt.WindowState.WindowMaximized & self.windowState())
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
        # with open('common/assets/settings.ini', 'w') as configfile: 
        #     self.config.write(configfile)
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
        

def main():
    app = QtWidgets.QApplication([sys.argv])
    #QtGui.QFontDatabase.addApplicationFont("assets/fonts/Roboto.ttf")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()

#articles on threading - need to implement both this and dataframe column assigning instead of creating a billion loops
#https://www.pythonguis.com/faq/real-time-change-of-widgets/
#https://nikolak.com/pyqt-threading-tutorial/