import pyodbc
import pandas as pd
import numpy as np
import os
import time
import common.AGS4_package_edit as AGS4 # had to edit this to concat linebreaks - credits to python_ags4, asitha-sena, https://gitlab.com/ags-data-format-wg/ags-python-library
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMessageBox, QWidget
from PyQt5.QtCore import pyqtSignal
from PyQt5 import QtCore, QtGui, uic
from rich import print as rprint
import warnings
warnings.filterwarnings("ignore", category=UserWarning)

class GintHandler(QWidget):
    _disable = pyqtSignal()
    _enable = pyqtSignal()
    _update_text = pyqtSignal(str)
    _gint_error_flag = pyqtSignal(bool)

    def __init__(self):
        super(GintHandler, self).__init__()
        self.gint_location: str = None
        self.gint_spec: pd.DataFrame = None
        self.config: object = None

    
    def get_gint(self):
        self._disable.emit()
        self._update_text.emit('''Getting gINT, please wait...
''')

        if not self.config.get('LastFolder','dir') == "":
            self.gint_location = QtWidgets.QFileDialog.getOpenFileNames(self,'Open gINT Project', self.config.get('LastFolder','dir'), '*.gpj')
        else:
            self.gint_location = QtWidgets.QFileDialog.getOpenFileNames(self,'Open gINT Project', os.getcwd(), '*.gpj')

        if len(self.gint_location[0]) == 0:
            self._enable.emit()
            return
        else:
            self.gint_location = self.gint_location[0][0]
            self.query_spec()

        
    def query_spec(self):
        driver_check = [x for x in pyodbc.drivers() if 'Access Driver' in x]
        if len(driver_check) == 0:
            msgBox = QMessageBox()
            msgBox.setIcon(QMessageBox.Information)
            msgBox.setText('''64-bit Access Driver not found.
Please contact Infinity to install this driver.''')
            msgBox.setWindowTitle("Access Driver Error")
            msgBox.setStandardButtons(QMessageBox.Ok)
            msgBox.exec()
            self._enable.emit()
            return self._gint_error_flag.emit(True)
        try:
            conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+self.gint_location+';')
            query = "SELECT * FROM SPEC"
            self.gint_spec = pd.read_sql(query, conn)
            self._gint_error_flag.emit(False)
        except Exception as e:
            print(e)
            print("Uhh.... either that's the wrong gINT, or something went wrong.")
            self._update_text.emit('''Uhh.... something went wrong.
''')
            self._enable.emit()
        
    

class AGSHandler(QWidget):
    _disable = pyqtSignal()
    _enable = pyqtSignal()
    _update_text = pyqtSignal(str)
    _open = pyqtSignal(bool)
    _enable_error_export = pyqtSignal(bool)
    _enable_results_export = pyqtSignal(bool)
    _set_model = pyqtSignal(pd.DataFrame)
    _coin = pyqtSignal()
    _table_setup = pyqtSignal()
    _progress_max = pyqtSignal(int)
    _progress_current = pyqtSignal(int)

    def __init__(self):
        super(AGSHandler, self).__init__()
        self.gint_location: str = None
        self.gint_spec: pd.DataFrame = None
        self.config: object = None
        self.tables: dict = None
        self.headings: dict = None
        self.ags_tables: list = []
        self.result_list: list = []
        self.error_list: list = []
        self.results_with_samp_and_type: pd.DataFrame = None
        self.temp_file_name: str = ''

        self.result_tables = ['SAMP','SPEC','TRIG','TRIT','LNMC','LDEN','GRAG','GRAT',
        'CONG','CONS','CODG','CODT','LDYN','LLPL','LPDN','LPEN','LRES','LTCH','LTHC',
        'LVAN','LHVN','RELD','SHBG','SHBT','TREG','TRET','DSSG','DSST','IRSG','IRST',
        'GCHM','ERES','RESG','REST','RESV','RESD','TORV','PTST','RPLT','RCAG','RDEN',
        'RUCS', 'RWCO', 'IRSV', 'TXTG', 'TXTT', 'SSTG', 'SSTT']

        self.core_tables = ["TRAN","PROJ","UNIT","ABBR","TYPE","DICT","LOCA"]


    def load_ags_file(self):
        self._disable.emit()

        self._update_text.emit('''Please insert AGS file.
''')

        if not self.config.get('LastFolder','dir') == "":
            self.file_location = QtWidgets.QFileDialog.getOpenFileNames(self,'Please insert AGS file...', self.config.get('LastFolder','dir'), '*.ags')
        else:
            self.file_location = QtWidgets.QFileDialog.getOpenFileNames(self,'Please insert AGS file...', os.getcwd(), '*.ags')
        if not len(self.file_location[0]) == 0:
            self.file_location = self.file_location[0][0]
            last_dir = str(os.path.dirname(self.file_location))
            self.config.set('LastFolder','dir',last_dir)
            with open('common/assets/settings.ini', 'w') as configfile: 
                self.config.write(configfile)
            self._update_text.emit('''AGS file loaded.
''')
        else:
            self._update_text.emit('''No AGS file selected!
Please select an AGS with "Open File..."''')
            print("No AGS file selected! Please select an AGS with 'Open File...'")
            self._disable.emit()
            self._open.emit(True)
            return
        
    def ags_tables_from_file(self):
        try:
            self.tables, self.headings = AGS4.AGS4_to_dataframe(self.file_location)
        except:
            print("Uh, something went wrong. Was that an AGS file? Send help.")
            self._open.emit(True)
        finally:
            rprint(f"[green]AGS file loaded: [/green][white][i][b]{self.file_location}[/b][/i][/white]")
            self._coin.emit()
            self._enable.emit()
            return self.tables, self.headings

    def get_ags_tables(self):
        self.ags_table_reset()

        for table in self.result_tables:
            if table in list(self.tables):
                self.ags_tables.append(table)
        return self.ags_tables
    
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
                rprint(f"[bold]{str(table)}[/bold] table [red]deleted.[/red]")
            try:
                if table == 'SAMP' or table == 'SPEC':
                    del self.tables[table]
                    rprint(f"[bold]{str(table)}[/bold] table [red]deleted.[/red]")
            except:
                pass
        self._update_text.emit('''Deleted non-lab testing tables,
select a lab to match to.''')
    
            
    def check_ags(self):
        self.error_list = []
        self._disable.emit()
        self._update_text.emit('''Checking AGS for errors...
''')
        try:
            if self.file_location == '':
                self._update_text.emit('''No AGS file selected!
Please select an AGS with "Open File..."''')
                print("No AGS file selected! Please select an AGS with 'Open File...'")
            else:
                try:
                    #errors = AGS4.check_file(self.file_location)
                    '''need to get the latest data from self.tables to check errors, but dataframe_to_AGS4 method returns a file... so create a temp file to delete later'''
                    AGS4.dataframe_to_AGS4(self.tables, self.tables, f'{os.getcwd()}\\_temp_.ags')
                    errors = AGS4.check_file(f'{os.getcwd()}\\_temp_.ags')
                except Exception as e:
                    print(e)
                    
        except ValueError as e:
            print(f'AGS Checker ended unexpectedly: {e}')
            try:
                if os.path.isfile(f'{os.getcwd()}\\_temp_.ags'):
                    os.remove(f'{os.getcwd()}\\_temp_.ags')
            except Exception as e:
                print(e)
            return
        
        for rule, items in errors.items():
            if rule == 'Metadata':
                print('Metadata')
                for msg in items:
                    print(f"{msg['line']}: {msg['desc']}")
                continue
                    
            for error in items:
                print(f"Error in line: {error['line']}, group: {error['group']}, description: {error['desc']}")
                self.error_list.append(f"Error in line: {error['line']}, group: {error['group']}, description: {error['desc']}")

        if errors:
            self._enable_error_export.emit(True)

            try:    
                if os.path.isfile(f'{os.getcwd()}\\_temp_.ags'):
                    os.remove(f'{os.getcwd()}\\_temp_.ags')
            except Exception as e:
                print(e)
            self._enable.emit()

            if self.error_list == []:
                self.error_list.append("No errors found. Yay.")
                print("No errors found. Yay.")
                self._update_text.emit("""AGS file contains no errors!
""")        
            else:
                self._update_text.emit('''Error(s) found, check output or click 'Export Error Log'.
''')

            err_df = pd.DataFrame.from_dict(self.error_list)
            self._set_model.emit(err_df)


    def export_errors(self):
        self._disable.emit()
        
        if not self.config.get('LastFolder','dir') == "":
            self.log_path = QtWidgets.QFileDialog.getSaveFileName(self,'Save error log as...', self.config.get('LastFolder','dir'), '*.txt')
        else:
            self.log_path = QtWidgets.QFileDialog.getSaveFileName(self,'Save error log as...', os.getcwd(), '*.txt')
        try:
            self.log_path = self.log_path[0]
            with open(self.log_path, "w") as f:
                for item in self.error_list:
                    f.write("%s\n" % item)
            self._enable.emit()
            self._enable_error_export.emit(True)
        except:
            self._enable.emit()
            self._enable_error_export.emit(True)
            return

        print(f"Error log exported to:  + {str(self.log_path)}")
        self._enable.emit()
        self._enable_error_export.emit(True)


    def save_ags(self):
        self._disable.emit()

        if not self.config.get('LastFolder','dir') == "":
            newFileName = QtWidgets.QFileDialog.getSaveFileName(self,'Save AGS file as...', self.config.get('LastFolder','dir'), '*.ags')
        else:
            newFileName = QtWidgets.QFileDialog.getSaveFileName(self,'Save AGS file as...', os.getcwd(), '*.ags')
        try:
            newFileName = newFileName[0]

            rprint(f"""[cyan]------------------------------------------------------
Saving AGS file...
------------------------------------------------------[/cyan]""")
            
            AGS4.dataframe_to_AGS4(self.tables, self.tables, newFileName)
            self._update_text.emit('''AGS saved.
''')
        
            rprint(f"""[green][bold]AGS saved:[/bold][/green] [white][i]{newFileName}""")
            self._enable.emit()
        except:
            self._enable.emit()
            return
        
    def count_lab_results(self):
        self._disable.emit()

        self.results_with_samp_and_type = pd.DataFrame()

        lab_tables = ['TRIG','LNMC','LDEN','GRAT','CONG','LDYN','LLPL','LPDN','LPEN',
        'LRES','LTCH','LVAN','RELD','SHBG','TREG','DSSG','IRSG','PTST','GCHM','RESG',
        'ERES','RCAG','RDEN','RUCS','RPLT','LHVN','TXTG', 'SSTG'
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
                    test_type_df = pd.DataFrame
                    location = list(self.tables[table]['LOCA_ID'][2:])
                    samp_id = list(self.tables[table]['SAMP_ID'][2:])
                    samp_ref = list(self.tables[table]['SPEC_REF'][2:])
                    samp_depth = list(self.tables[table]['SPEC_DPTH'][2:])

                    lab_field = [col for col in self.tables[table].columns if 'LAB' in col]
                    if not lab_field == []:
                        lab_field = str(lab_field[0])
                    lab = self.tables[table][lab_field].iloc[2:]

                    if 'GCHM' in table:
                        test_type_df = pd.DataFrame.from_dict(list(self.tables[table]['GCHM_CODE'][2:]))
                    elif 'TRIG' in table:
                        test_type_df = pd.DataFrame.from_dict(list(self.tables[table]['TRIG_COND'][2:]))
                    elif 'CONG' in table:
                        test_type_df = pd.DataFrame.from_dict(list(self.tables[table]['CONG_COND'][2:]))
                    elif 'TREG' in table:
                        test_type_df = pd.DataFrame.from_dict(list(self.tables[table]['TREG_TYPE'][2:]))
                    elif 'ERES' in table:
                        test_type_df = pd.DataFrame.from_dict(list(self.tables[table]['ERES_TNAM'][2:]))
                    elif 'TXTG' in table:
                        test_type_df = pd.DataFrame.from_dict(list(self.tables[table]['TXTG_TYPE'][2:]))
                    elif 'SSTG' in table:
                        test_type_df = pd.DataFrame.from_dict(list(self.tables[table]['SSTG_TYPE'][2:]))
                    elif 'GRAT'in table:
                        test_type = list(self.tables[table]['GRAT_TYPE'][2:])
                        #turning into a dataframe to drop all the duplcicates per test type per sample, using the test type from that list as the count
                        samp_with_table = list(zip(location,samp_id,samp_ref,samp_depth,test_type))
                        result_table = pd.DataFrame.from_dict(samp_with_table)
                        result_table.drop_duplicates(inplace=True)
                        result_table.columns = ['POINT','ID','REF','DEPTH','TYPE']
                        tt = result_table['TYPE'].to_list()
                        test_type_df = pd.DataFrame.from_dict(tt)
                    if not type(lab) == list and lab.empty:
                        lab = list(samp_id)
                        for x in range(0,len(lab)):
                            lab[x] = ""

                    if test_type_df.empty:
                        test_type = []
                    else:
                        test_type = list(test_type_df[0])

                    samples = list(zip(location,samp_id,samp_ref,samp_depth,test_type,lab))
                    table_results = pd.DataFrame.from_dict(samples)
                    if table == 'GRAT':
                        table_results.drop_duplicates(inplace=True)
        
                    if not len(test_type) == 0:
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
                        if not lab_count_none == 0:
                            if "GRAT" in table:
                                valcount = len(list(set(list(zip(samp_id,samp_depth)))))
                                #valcount = table_results.shape[0] - 1
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

        result_list = self.result_list.to_string(col_space=10,justify="center",index=None, header=None)
        self._set_model.emit(self.result_list)
        print(result_list)

        self._update_text.emit('''Results list ready to export.
''')
        self._enable_results_export.emit(True)
        self._enable.emit()

        
    def export_results(self):
        self._disable.emit()

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
            self._enable.emit()
            self._enable_results_export.emit(True)
        except:
            self._enable.emit()
            self._enable_results_export.emit(True)
            return
    

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

            self._update_text.emit('''CPT Only export ready.
Click "Save AGS file"''')
            print("CPT Data export ready. Click 'Save AGS file'.")

        else:
            self._update_text.emit('''Could not find any CPT tables.
Check the AGS with "View data".''')
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

            self._update_text.emit('''Lab Data & GEOL export ready.
Click "Save AGS file"''')
            print("Lab Data & GEOL export ready. Click 'Save AGS file'.")
            self.ags_table_reset()

        else:
            self._update_text.emit('''Could not find any Lab or GEOL tables.
Check the AGS with "View data".''')
            print("No Lab or GEOL groups found - did this AGS contain CPT data? Check the data with 'View data'.")
            self.ags_table_reset()

    def convert_excel(self):
        try:
            fname = QtWidgets.QFileDialog.getSaveFileName(self, "Save AGS as excel...", os.path.dirname(self.file_location), "Excel file *.xlsx;")
        except:
            fname = QtWidgets.QFileDialog.getSaveFileName(self, "Save AGS as excel...", os.getcwd(), "Excel file *.xlsx;")
        
        if fname[0] == '':
            return

        final_dataframes = [(k,v) for (k,v) in self.tables.items() if not v.empty]
        final_dataframes = dict(final_dataframes)
        empty_dataframes = [k for (k,v) in self.tables.items() if v.empty]

        progress = 0
        progress_total = (len(self.tables.keys())) * 100
        self._progress_max.emit(progress_total)
        self._progress_current.emit(progress)   

        rprint(f"""[cyan]------------------------------------------------------
Saving AGS to excel file...
------------------------------------------------------[/cyan]""")

        #create the excel file with the first dataframe from dict, so pd.excelwriter can be called (can only be used on existing excel workbook to append more sheets)
        if not len(final_dataframes.keys()) < 1:
            next(iter(final_dataframes.values())).to_excel(f"{fname[0]}", sheet_name=(f"{next(iter(final_dataframes))}"), index=None, index_label=None)
            final_writer = pd.ExcelWriter(f"{fname[0]}", engine="openpyxl", mode="a", if_sheet_exists="replace")
        else:
            print(f"All selected tables are empty! Please select others. Tables selected: {empty_dataframes}")
            return

        #for every key (table name) and value (table data) in the AGS, append to excel sheet and update progress bar, saving only at the end for performance
        for (k,v) in final_dataframes.items():
            progress += 100
            self._progress_current.emit(progress)  
            rprint(f"[green]Writing [bold]{k}[/bold] to excel...[green]")
            self._update_text.emit(f'''Writing {k} to excel...
''')
            v.to_excel(final_writer, sheet_name=(f"{str(k)}"), index=None, index_label=None)
            time.sleep(0.01)
        final_writer.close()

        rprint(f"""[green][bold]AGS saved as Excel file:[/bold][/green] [white][i]{fname[0]}""")
        self._update_text.emit(f'''AGS saved as Excel file.
''')
        

class DataframeProcessor:
    def __init__(self):
        pass

    def search_dataframe(self, value, df, table):
        data = df

        if isinstance(value, str):
            mask = data.astype(str) == value
        else:
            mask = data == value

        search_result = np.where(mask)

        if search_result[0].size == 0:
            print("Value not found")
            return None, None
        else:
            first_row, first_col = search_result[0][0], search_result[1][0]
            match_value = data.iloc[first_row, first_col]
            print(f"Value: {match_value}\nRow: {first_row}\nColumn: {first_col}")
            return first_row, first_col

    def fill_df(self, df, table):
        df.replace([None, ''], np.nan, inplace=True)

        def convert_float_ints(col): # need vals as nan to use ffill(), but adding nan to cols with ints and empty strings turns col to float - this will convert floats to int in those cols
            if all(isinstance(x, float) and x.is_integer() for x in col if isinstance(x, float)):
                return col.apply(lambda x: int(x) if isinstance(x, float) and (x).is_integer() else x)
            return col
        
        for column in df.columns:
            non_null_count = df[column].notna().sum()
            
            # if non_null_count == 1:
            #     df[column].ffill(inplace=True)
            # else:
            df[column].fillna("", inplace=True)
            df[column] = convert_float_ints(df[column])
            
        table.model().df = df
        table.model().layoutChanged.emit()

    def sample_fill(self, df, table):
        columns = self.get_current_columns_multiple(table)

        if not columns:
            return

        df[columns] = df[columns].replace([None, '', np.nan], np.nan).ffill()

        table.model().df = df
        table.model().layoutChanged.emit()

    def replace_df(self, df, table, find_text, replace_text, df_radio, col_radio, cell_radio):
        if not cell_radio:
            df_head = df.iloc[:2].copy()
            df = df.iloc[2:].copy()

        find_type = self.validate_text_type(find_text)
        replace_type = self.validate_text_type(replace_text)

        column_name = self.get_current_columns_multiple(table)
        cell_iloc = self.get_current_cell_iloc(table)

        def replace_value(x):
            if pd.isna(x) and find_text.lower() == "nan":
                return replace_text
            if isinstance(x, list):
                return x
            try:
                return str(replace_type) if float(x) == find_type else x
            except ValueError:
                return x.replace(find_text, replace_text) if isinstance(x, str) else x

        if df_radio:
            df = df.applymap(replace_value)
        elif col_radio:
            if column_name:
                for col in column_name:
                    df[col] = df[col].apply(replace_value)
        elif cell_radio:
            if cell_iloc:
                for (row, col) in cell_iloc:
                    if pd.isna(df.iat[row, col]):
                        df.iat[row, col] = replace_text
                    else:
                        df.iat[row, col] = replace_value(df.iat[row, col])

        if not cell_radio:
            df = pd.concat([df_head, df], ignore_index=True)
        table.model().df = df
        table.model().layoutChanged.emit()

    def format_df(self, df, table, decimal_places, df_radio, col_radio, cell_radio):
        if not cell_radio:
            df_head = df.iloc[:2].copy()
            df = df.iloc[2:].copy()
    
        def round_value(x):
            if pd.isna(x):
                return ""
            try:
                if decimal_places == 0:
                    return int(round(float(x)))
                else:
                    return round(float(x), decimal_places)
            except (ValueError, TypeError):
                return x

        column_name = self.get_current_columns_multiple(table)
        cell_iloc = self.get_current_cell_iloc(table)

        if df_radio:
            df = df.applymap(round_value)
        elif col_radio:
            if column_name:
                for col in column_name:
                    df[col] = df[col].apply(round_value)
        elif cell_radio:
            if cell_iloc:
                for (row, col) in cell_iloc:
                    df.iat[row, col] = round_value(df.iat[row, col])

        if not cell_radio:
            df = pd.concat([df_head, df], ignore_index=True)
        table.model().df = df
        table.model().layoutChanged.emit()

    def split_df(self, df, table, delimiter):
        column = self.get_current_column(table)

        if not column:
            return

        split_columns = df[column].str.split(delimiter, expand=True)

        if split_columns.shape[1] > 1:
            col_index = df.columns.get_loc(column) + 1
            for i in range(split_columns.shape[1]):
                new_col_name = f"{column}_{i+1}"
                df.insert(col_index + i, new_col_name, split_columns[i])
            df.drop(column, axis=1, inplace=True)
        else:
            print("Couldn't split on selected delimiter.")

        table.model().df = df
        table.model().layoutChanged.emit()

    def case_df(self, df, table, case, df_radio, col_radio, cell_radio):
        if not cell_radio:
            df_head = df.iloc[:2].copy()
            df = df.iloc[2:].copy()
    
        column_names = self.get_current_columns_multiple(table)
        cell_iloc = self.get_current_cell_iloc(table)

        if df_radio:
            for col in df.columns:
                if all(isinstance(x, str) for x in df[col] if pd.notna(x)):
                    if case == "Upper Case":
                        df[col] = df[col].apply(lambda x: x.upper() if isinstance(x, str) else x)
                    elif case == "Lower Case":
                        df[col] = df[col].apply(lambda x: x.lower() if isinstance(x, str) else x)
                    elif case == "Capitalise":
                        df[col] = df[col].apply(lambda x: x.capitalize() if isinstance(x, str) else x)
        elif col_radio:
            if column_names:
                for col in column_names:
                    if all(isinstance(x, str) for x in df[col] if pd.notna(x)):
                        if case == "Upper Case":
                            df[col] = df[col].apply(lambda x: x.upper() if isinstance(x, str) else x)
                        elif case == "Lower Case":
                            df[col] = df[col].apply(lambda x: x.lower() if isinstance(x, str) else x)
                        elif case == "Capitalise":
                            df[col] = df[col].apply(lambda x: x.capitalize() if isinstance(x, str) else x)
        elif cell_radio:
            if cell_iloc:
                for (row, col) in cell_iloc:
                    if isinstance(df.iat[row, col], str): # this is shit and needs to change - numbers could be stored as strings
                        if case == "Upper Case":
                            df.iat[row, col] = df.iat[row, col].upper() if isinstance(df.iat[row, col], str) else df.iat[row, col]
                        elif case == "Lower Case":
                            df.iat[row, col] = df.iat[row, col].lower() if isinstance(df.iat[row, col], str) else df.iat[row, col]
                        elif case == "Capitalise":
                            df.iat[row, col] = df.iat[row, col].capitalize() if isinstance(df.iat[row, col], str) else df.iat[row, col]

        if not cell_radio:
            df = pd.concat([df_head, df], ignore_index=True)
        table.model().df = df
        table.model().layoutChanged.emit()

    def calc_df(self, df, table, calculation, value, df_radio, col_radio, cell_radio):
        if not cell_radio:
            df_head = df.iloc[:2].copy()
            df = df.iloc[2:].copy()
    
        column_names = self.get_current_columns_multiple(table)
        cell_iloc = self.get_current_cell_iloc(table)

        def apply_calculation(column):
            if calculation == "Multiply":
                return column * value
            elif calculation == "Divide":
                return column / value if value != 0 else None
            elif calculation == "Add":
                return column + value
            elif calculation == "Subtract":
                return column - value
            elif calculation == "Average":
                return column

        def convert_and_apply(column):
            numeric_col = pd.to_numeric(column, errors='coerce')
            if numeric_col.notna().any():
                return apply_calculation(numeric_col)
            return column

        if df_radio:
            for col in df.columns:
                df[col] = convert_and_apply(df[col])

        elif col_radio:
            if column_names:
                for col in column_names:
                    df[col] = convert_and_apply(df[col])
                if calculation == "Average":
                    avg_type = self.display_combo_popup(items=['Rows', 'Columns'], title="", label="Average on:", win_title="Average Type")
                    if not avg_type:
                        return
                    avg_col, ok = QtWidgets.QInputDialog.getText(None, 'Column Name', 'Average Column Name:')
                    if not ok:
                        avg_col = f"Averages for ({', '.join(column_names)})"

                    if avg_type == "Rows":
                        df[avg_col] = df[column_names].mean(axis=1)
                    elif avg_type == "Columns":
                        if all(df[col].dtype.kind in 'biufc' for col in column_names):
                            df[avg_col] = np.nanmean(df[column_names].values.flatten()) # ignore nan for avg of all cols
                            
        elif cell_radio:
            if cell_iloc:
                values = []
                for (row, col) in cell_iloc:
                    converted_value = pd.to_numeric(df.iat[row, col], errors='coerce')
                    if pd.notna(converted_value):
                        if calculation != "Average":
                            df.iat[row, col] = apply_calculation(converted_value)
                        else:
                            values.append(converted_value)
                if calculation == "Average" and values:
                    overall_average = sum(values) / len(values)
                    involved_rows = set(row for row, col in cell_iloc)
                    avg_col, ok = QtWidgets.QInputDialog.getText(None, 'Column Name', 'Average Column Name:')
                    if not ok:
                        avg_col = f"Cell Averages for ({', '.join(column_names)})"
                    for row in involved_rows:
                        df.at[row, avg_col] = overall_average

        if not cell_radio:
            df = pd.concat([df_head, df], ignore_index=True)
        table.model().df = df
        table.model().layoutChanged.emit()

    #//==============================================================================
    # HELPER FUNCTIONS
    #//==============================================================================

    def validate_text_type(self, text):
        text = text.strip()
        try:
            float_value = float(text)
            if float_value.is_integer():
                return int(float_value)
            return float_value
        except ValueError:
            return text
        

    def get_current_column(self, table):
        index = table.selectionModel().currentIndex()
        if index.isValid():
            column = index.column()
            column_name = table.model().headerData(column, QtCore.Qt.Orientation.Horizontal, QtCore.Qt.ItemDataRole.DisplayRole)
            return column_name
        return False
    

    def get_current_columns_multiple(self, table):
        index = table.selectionModel().currentIndex()
        if index.isValid():
            unique_columns = set(index.column() for index in table.selectionModel().selectedIndexes())
            column_names = [table.model().headerData(column, QtCore.Qt.Orientation.Horizontal, QtCore.Qt.ItemDataRole.DisplayRole) for column in unique_columns]
            return column_names
        return False


    def get_current_cell_iloc(self, table):
        indexes = table.selectionModel().selectedIndexes()
        if not indexes:
            return False
        iloc = list(set((index.row(), index.column()) for index in indexes))
        return iloc
    

    def get_current_files(self):
        files = []
        for index in range(self.file_combo.count()):
            files.append(self.file_combo.itemText(index))
        return files
    

    def get_delimiter(self):
        if self.delimit_combo.currentText() == "Comma":
            return ","
        elif self.delimit_combo.currentText() == "Hyphen":
            return "-"
        elif self.delimit_combo.currentText() == "Underscore":
            return "_"
        elif self.delimit_combo.currentText() == "Space":
            return " "
        elif self.delimit_combo.currentText() == "Decimal":
            return "."
        elif self.delimit_combo.currentText() == "Colon":
            return ":"
        elif self.delimit_combo.currentText() == "Semi-colon":
            return ";"
        
    def display_combo_popup(self, items, title, label, win_title):
        combo_popup = ComboPopup(items=items, text_title=title, text_label=label, win_title=win_title)
        combo_popup.show()
        loop = QtCore.QEventLoop()
        combo_popup.finished.connect(loop.quit)
        loop.exec()
        return combo_popup._value if combo_popup._value else None
    

class ComboPopup(QtWidgets.QWidget):
    '''Popup with a ComboBox, Push Button and titles for labels and window so it can be re-used'''
    
    finished = QtCore.pyqtSignal(str)

    def __init__(self, items, text_title, text_label, win_title):
        super().__init__()
        uic.loadUi("common/assets/ui/combo_popup.ui", self)
        self.text_title.setVisible(False)
        self.resize(self.width(), self.minimumSizeHint().height())
        self.setWindowIcon(QtGui.QIcon("common/images/geo.ico"))
        self.setWindowFlags(QtCore.Qt.WindowType.WindowStaysOnTopHint) # Force top level
        self.setWindowTitle(win_title)
        self._text_label = text_label
        self._text_title = text_title
        self._items = items
        self._value: str = ''
        self.get_value.clicked.connect(self.return_value)
        self.set_label()
        self.set_title()
        self.populate_combo()

    def set_label(self):
        self.text_label.setText(self._text_label)

    def set_title(self):
        pass
        #self.text_title.setText(self._text_title)

    def populate_combo(self):
        for val in self._items:
            if not val == '':
                self.popup_combo.addItem(val)

    def return_value(self):
        self._value = self.popup_combo.currentText()
        self.close()

    def closeEvent(self, a0: QtGui.QCloseEvent) -> None:
        return self.finished.emit(self._value)
    