import pyodbc
import pandas as pd
import os
import time
from python_ags4 import AGS4
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMessageBox, QWidget
from PyQt5.QtCore import pyqtSignal
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
            print(f"AGS file loaded: {self.file_location}")
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
                print(f"{str(table)} table deleted.")
            try:
                if table == 'SAMP' or table == 'SPEC':
                    del self.tables[table]
                    print(f"{str(table)} table deleted.")
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
            AGS4.dataframe_to_AGS4(self.tables, self.tables, newFileName)
            print('Done.')
            self._update_text.emit('''AGS saved.
''')
        
            print(f"""AGS saved: {newFileName}""")
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

        print(f"""------------------------------------------------------
Saving AGS to excel file...
------------------------------------------------------""")

        #create the excel file with the first dataframe from dict, so pd.excelwriter can be called (can only be used on existing excel workbook to append more sheets)
        if not len(final_dataframes.keys()) < 1:
            next(iter(final_dataframes.values())).to_excel(f"{fname[0]}", sheet_name=(f"{next(iter(final_dataframes))}"), index=None, index_label=None)
            final_writer = pd.ExcelWriter(f"{fname[0]}", engine="openpyxl", mode="a", if_sheet_exists="replace")
        else:
            print(f"All selected tables are empty! Please select others. Tables selected: {empty_dataframes}")
            return

        #for every key (table name) and value (table data) in the AGS, append to excel sheet and update progress bar, saving only at the end for performance
        for (k,v) in final_dataframes.items():
            print(f"Writing {k} to excel...")
            v.to_excel(final_writer, sheet_name=(f"{str(k)}"), index=None, index_label=None)
            time.sleep(0.01)
        final_writer.close()

        print(f"""AGS saved as Excel file: {fname[0]}""")