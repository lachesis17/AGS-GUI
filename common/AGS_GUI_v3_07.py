import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from python_ags4 import AGS4
from pandasgui import show
from statistics import mean
import pandas as pd
import customtkinter as ct
from customtkinter.windows.widgets.scaling import CTkScalingBaseClass
import pyodbc
import warnings
#from common.assets.splash import SplashWin
warnings.filterwarnings("ignore")

def main():
    ct.deactivate_automatic_dpi_awareness()

    class Application(ct.CTkFrame):

        ct.set_appearance_mode("system")
        ct.set_default_color_theme("dark-blue")

        def __init__(self):
            super(Application, self).__init__(window)

            CTkScalingBaseClass.destroy(window)
            CTkScalingBaseClass.destroy(self)
            ct.deactivate_automatic_dpi_awareness()

            #window.iconphoto(False, tk.PhotoImage(file='images/geo.png'))
            window.wm_iconbitmap(bitmap='common/images/geo.ico', default='common/images/geo.ico')
            window.lift()
            window.geometry('450x440+150+150')
            window.resizable(False,False)
            window.title("AGS GUI v3.07")

            self.botframe = ct.CTkFrame(window)
            self.botframe.pack(pady=(0,16), padx=8, side=tk.BOTTOM)
            
            self.button_open = ct.CTkButton(self, text="Open File...", command=self.get_ags_file, fg_color="#2b4768", 
            corner_radius=10, hover_color="#6bb7dd", text_color="#FFFFFF", text_color_disabled="#999999", font=("Tahoma",11), width=300)
            self.button_open.pack(pady=8, padx=8)

            self.pandas_gui = ct.CTkButton(self, text="View Data", command=self.start_pandasgui, 
            corner_radius=10, fg_color="#2b4768", hover_color="#6bb7dd", text_color="#FFFFFF", text_color_disabled="#999999", font=("Tahoma",11), width=200)
            self.pandas_gui.pack(pady=8)
            self.pandas_gui.configure(state=tk.DISABLED)

            self.button_save_ags = ct.CTkButton(self, text="Save AGS File", command=self.save_ags, 
            corner_radius=10, fg_color="#2b4768", hover_color="#6bb7dd", text_color="#FFFFFF", text_color_disabled="#999999", font=("Tahoma",11), width=200)
            self.button_save_ags.pack(pady=8)
            self.button_save_ags.configure(state=tk.DISABLED)

            self.button_count_results = ct.CTkButton(self, text="Count Lab Results", command=self.count_lab_results, 
            corner_radius=10, fg_color="#2b4768", hover_color="#6bb7dd", text_color="#FFFFFF", text_color_disabled="#999999", font=("Tahoma",11), width=150)
            self.button_count_results.pack(pady=8, padx=10, side=tk.LEFT)
            self.button_count_results.configure(state=tk.DISABLED)
            
            self.button_ags_checker = ct.CTkButton(self, text="Check AGS for Errors", command=self.check_ags, 
            corner_radius=10, fg_color="#2b4768", hover_color="#6bb7dd", text_color="#FFFFFF", text_color_disabled="#999999", font=("Tahoma",11), width=150)
            self.button_ags_checker.pack(pady=8, padx=10, side=tk.RIGHT)
            self.button_ags_checker.configure(state=tk.DISABLED)

            self.button_del_tbl = ct.CTkButton(self.botframe, text='''Delete Non-Result Tables for gINT Import''', command=self.del_non_lab_tables, 
            corner_radius=10, fg_color="#2b4768", hover_color="#6bb7dd", text_color="#FFFFFF", text_color_disabled="#999999", font=("Tahoma",11), width=300)
            self.button_del_tbl.pack(pady=8, side=tk.TOP)
            self.button_del_tbl.configure(state=tk.DISABLED)

            self.selected_lab = ct.StringVar(value="Select a Lab")

            self.lab_select = ct.CTkOptionMenu(master=self.botframe, variable=self.selected_lab, values=["GM Lab","GM Lab PEZ","DETS", "DETS PEZ", "Structural Soils","PSL","Geolabs","Geolabs (50HZ Fugro)"],
            corner_radius=10, fg_color="#2b4768", text_color="#FFFFFF", text_color_disabled="#999999", font=("Tahoma",11), width=200)
            self.lab_select.pack(pady=8, side=tk.TOP)
            self.lab_select.configure(state=tk.DISABLED)

            self.button_match_lab = ct.CTkButton(self.botframe, text='''Match Lab AGS to gINT''', command=self.select_lab_match, 
            corner_radius=10, fg_color="#2b4768", hover_color="#6bb7dd", text_color="#FFFFFF", text_color_disabled="#999999", font=("Tahoma",11), width=200)
            self.button_match_lab.pack(pady=8, side=tk.TOP)
            self.button_match_lab.configure(state=tk.DISABLED)

            self.button_cpt_only = ct.CTkButton(self.botframe, text='''CPT Only Data Export''', command=self.del_non_cpt_tables, 
            corner_radius=10, fg_color="#2b4768", hover_color="#6bb7dd", text_color="#FFFFFF", text_color_disabled="#999999", font=("Tahoma",11), width=150)
            self.button_cpt_only.pack(pady=8, padx=10, side=tk.RIGHT)
            self.button_cpt_only.configure(state=tk.DISABLED)

            self.button_lab_only = ct.CTkButton(self.botframe, text='''Lab Only Data Export''', command=self.export_lab_only, 
            corner_radius=10, fg_color="#2b4768", hover_color="#6bb7dd", text_color="#FFFFFF", text_color_disabled="#999999", font=("Tahoma",11), width=150)
            self.button_lab_only.pack(pady=8, padx=10, side=tk.LEFT)
            self.button_lab_only.configure(state=tk.DISABLED)
        
            self.text = tk.StringVar()
            self.text.set('''Please insert AGS file.
    ''')

            self.greeting = ct.CTkLabel(window, textvariable=self.text, font=("Tahoma",11))
            self.greeting.pack(pady=(24,8))

            self.pack(padx=12)

            '''Empty variables created here to be used later. Also used to check previous actions, by checking if they're still empty in later functions'''
            self.temp_file_name = ''
            self.tables = None
            self.headings = None
            self.gui = None
            self.box = False   
            self.result_list = []
            self.error_list = []
            self.ags_tables = []
            self.export = False
            self.results_with_samp_and_type = ""

            self.core_tables = ["TRAN","PROJ","UNIT","ABBR","TYPE","DICT","LOCA"]

            self.result_tables = ['SAMP','SPEC','TRIG','TRIT','LNMC','LDEN','GRAG','GRAT',
            'CONG','CONS','CODG','CODT','LDYN','LLPL','LPDN','LPEN','LRES','LTCH','LTHC',
            'LVAN','LHVN','RELD','SHBG','SHBT','TREG','TRET','DSSG','DSST','IRSG','IRST',
            'GCHM','ERES','RESG','REST','RESV','RESD','TORV','PTST','RPLT','RCAG','RDEN',
            'RUCS', 'RWCO', 'IRSV']


        def get_ags_file(self):
            self.disable_buttons()

            window.geometry('450x440')
            self.text.set('''Please insert AGS file.
    ''')
            if self.box == True:
                self.listbox.pack_forget()
                self.button_export_results.pack_forget()
                self.box = False
            if self.export == True:
                self.button_export_error.pack_forget()
                self.export = False

            self.file_location = filedialog.askopenfilename(filetypes=[('AGS Files', '*.ags')],title="Please insert AGS file...")
            
            '''If no file selected, '''
            if not self.file_location:
                self.text.set('''No AGS file selected!
    Please select an AGS with "Open File..."''')
                print("No AGS file selected! Please select an AGS with 'Open File...'")
                self.button_open.configure(state=tk.NORMAL)
                return
            else:
                self.text.set('''AGS file loaded.
    ''')
                window.update()

            try:
                self.tables, self.headings = AGS4.AGS4_to_dataframe(self.file_location)
            except:
                print("Uh, something went wrong. Was that an AGS file? Send help.")
                self.button_open.configure(state=tk.NORMAL)
            finally:
                print(f"AGS file loaded: {self.file_location}")
                self.enable_buttons()

        def count_lab_results(self):
            self.disable_buttons()

            if self.export == True:
                self.button_export_error.pack_forget()
                self.export = False

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
                    #try:
                        print(table)
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
                        
                        lab = self.tables[table][lab_field]
                        if not lab.empty:
                            lablist = list(lab)
                            lablist.pop(0)
                            lablist.pop(0)
                            lab = pd.DataFrame.from_dict(lablist)
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
                            self.tables[table]["LAB"] = self.tables[table]['GRAT_TYPE']
                            lablist = self.tables[table]["LAB"]
                            lablist = lablist.iloc[2:]
                            lablist = list(lablist)
                            samp_with_table = list(zip(location,samp_id,samp_ref,samp_depth,test_type,lablist))
                            print(samp_with_table)
                            samp_with_table.pop(0)
                            samp_with_table.pop(0)
                            result_table = pd.DataFrame.from_dict(samp_with_table)
                            result_table.drop_duplicates(inplace=True)
                            result_table.columns = ['POINT','ID','REF','DEPTH','TYPE','LAB']
                            tt = result_table['TYPE'].to_list()
                            test_type_df = pd.DataFrame.from_dict(tt)

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
                        else:
                            count = str(len(samp_id))
                            sample = list(zip(location,samp_id,samp_ref,samp_depth))
                            table_results_2 = pd.DataFrame.from_dict(sample)
                            if table == 'RPLT':
                                table_results_2.drop_duplicates(inplace=True)
                                count = table_results_2.shape[0]
                            table_results_2.loc[-1] = [table,'','','']
                            table_results_2.index = table_results_2.index + 1
                            table_results_2.sort_index(inplace=True)
                            table_results = pd.concat([table_results, table_results_2])
                        type_list = []
                        type_list.append(str(table))
                        type_list.append(count)
                        all_results.append(type_list)
                        #print(str(table) + " - " + str(type_list))

                        self.results_with_samp_and_type = pd.concat([self.results_with_samp_and_type, table_results])

                    #except Exception as e:
                    #    error_tables.append(str(e))

            if error_tables != []:
                print(f"Table(s) not found:  {str(error_tables)}")

            self.result_list = pd.DataFrame.from_dict(all_results, orient='columns')

            if self.box == False:
                if self.result_list.empty:
                    df_list = ["Error: No laboratory test results found."]
                    empty_df = pd.DataFrame.from_dict(df_list)
                    self.result_list = empty_df
                self.listbox = scrolledtext.ScrolledText(self, height=10, font=("Tahoma",11))
                self.result_list.index.name = ' '
                window.geometry('775x620')
                self.listbox.tag_configure('tl', justify='left')
                self.listbox.insert('end', self.result_list, 'tl')
                self.listbox.delete(1.0,3.0)
                self.box = True

                self.button_export_results = ct.CTkButton(self, text="Export Results List", command=self.export_results, 
                corner_radius=10, fg_color="#2b4768", hover_color="#6bb7dd", text_color="#FFFFFF", text_color_disabled="#999999", font=("Tahoma",11), height=50, width=200)
                self.button_export_results.pack(pady=(8,8), side=tk.BOTTOM)
                self.listbox.pack(padx=20,pady=8, side=tk.BOTTOM)

                self.text.set('''Results list ready to export.
    ''')
            else:
                window.geometry('775x620')
                self.listbox.pack_forget()
                self.button_export_results.pack_forget()
                self.listbox.delete(1.0,100.0)
                if self.result_list.empty:
                    df_list = ["Error: No laboratory test results found."]
                    empty_df = pd.DataFrame.from_dict(df_list)
                    self.result_list = empty_df
                self.result_list.index.name = ' '
                self.listbox.tag_configure('tl', justify='left')
                self.listbox.insert('end', self.result_list, 'tl')
                self.listbox.delete(1.0,3.0)
                self.button_export_results.pack(pady=(8,8), side=tk.BOTTOM)
                self.listbox.pack(padx=20,pady=8, side=tk.BOTTOM)
                pass

            self.enable_buttons()
            
        def export_results(self):
            self.disable_buttons()

            self.results_with_samp_and_type.reset_index(inplace=True)
            self.results_with_samp_and_type.sort_index(inplace=True)

            if len(self.results_with_samp_and_type.columns) == 5:
                self.results_with_samp_and_type.loc[-1] = ['INDX','BH','ID','REF','DEPTH','LAB']
                self.results_with_samp_and_type.index = self.results_with_samp_and_type.index + 1
                self.results_with_samp_and_type.sort_index(inplace=True)
            else:
                self.results_with_samp_and_type.loc[-1] = ['INDX','BH','ID','REF','DEPTH','TYPE','LAB']
                self.results_with_samp_and_type.index = self.results_with_samp_and_type.index + 1
                self.results_with_samp_and_type.sort_index(inplace=True)

            self.path_directory = filedialog.asksaveasfilename(filetypes=[('CSV Files', '*.csv')],defaultextension="*.csv",title="Save results list as...")
            if not self.path_directory:
                self.enable_buttons()
                return
            all_result_count = self.path_directory[:-4] + "_result_count.csv"
            self.result_list.to_csv(all_result_count, index=False, index_label=False, header=None)
            print(f"File saved in:  + {str(all_result_count)}")
            all_result_filename = self.path_directory[:-4] + "_all_results.csv"
            self.results_with_samp_and_type.to_csv(all_result_filename, index=False,  header=None)	
            print(f"File saved in:  + {str(all_result_filename)}")
            self.enable_buttons()


        def start_pandasgui(self):
            self.disable_buttons()

            window.geometry('450x440')
            self.text.set('''PandasGUI loading, please wait...
    Close GUI to resume.''')
            window.update()
            if self.box == True:
                self.listbox.pack_forget()
                self.button_export_results.pack_forget()
                self.box = False
            if self.export == True:
                self.button_export_error.pack_forget()
                self.export = False
            window.update()
            
            try:
                self.gui = show(**self.tables)
            except:
                pass
            self.text.set('''You can now save the edited AGS.
    ''')
            window.update()
            updated_tables = self.gui.get_dataframes()
            self.tables = updated_tables
            for table in self.result_tables:
                if table in list(self.tables):
                    self.ags_tables.append(table)
            self.enable_buttons()
            
        def check_ags(self):
            self.disable_buttons()
            window.geometry('450x440')
            if self.box == True:
                self.listbox.pack_forget()
                self.button_export_results.pack_forget()
                self.box = False
            if self.export == True:
                self.button_export_error.pack_forget()
                self.export = False
            self.text.set('''Checking AGS for errors...
    ''')
            window.update()

            try:
                if not self.file_location == '' and self.temp_file_name == '':
                    errors = AGS4.check_file(self.file_location)
                elif not self.temp_file_name == '':
                    errors = AGS4.check_file(self.temp_file_name)
                else:
                    if self.file_location == '':
                        self.text.set('''No AGS file selected!
    Please select an AGS with "Open File..."''')
                        print("No AGS file selected! Please select an AGS with 'Open File...'")
                        self.button_ags_checker.configure(state=tk.DISABLED)
                    elif self.temp_file_name == '':
                        self.text.set('''No AGS file selected!
    Please select an AGS with "Open File..."''')
                        print("No AGS file selected! Please select an AGS with 'Open File...'")
                        self.button_ags_checker.configure(state=tk.DISABLED)
                    else:
                        errors = AGS4.check_file(self.file_location)
                        
            except ValueError as e:
                print(f'AGS Checker ended unexpectedly: {e}')
                return
            
            if not errors:
                print("No errors found. Yay.")
                self.text.set("AGS file contains no errors!")
                window.update()
                return
            
            for rule, items in errors.items():
                if rule == 'Metadata':
                    print('Metadata')
                    for msg in items:
                        print(f"{msg['line']}: {msg['desc']}")
                    continue
                        
                for error in items:
                    self.text.set('''Error(s) found, check output or click 'Export Error Log'.
    ''')
                    window.update()
                    print(f"Error in line: {error['line']}, group: {error['group']}, description: {error['desc']}")
                    self.error_list.append(f"Error in line: {error['line']}, group: {error['group']}, description: {error['desc']}")

            if errors:
                window.geometry('550x460')
                self.button_export_error = ct.CTkButton(self, text="Export Error Log", command=self.export_errors, 
                corner_radius=10, fg_color="#2b4768", hover_color="#6bb7dd", text_color="#FFFFFF", text_color_disabled="#999999", font=("Tahoma",11), height=50, width=200)
                self.button_export_error.pack(pady=(8,8), side=tk.BOTTOM)
                self.export = True
                self.text.set('''Error(s) found, check output or click 'Export Error Log'.
    ''')
                window.update()
            self.enable_buttons()

        def export_errors(self):
            self.disable_buttons()
            
            self.log_path = filedialog.asksaveasfilename(filetypes=[('Text Files', '*.txt')],defaultextension="*.txt",title="Save error log as...")
            if self.log_path == '': 
                self.enable_buttons()
                return
                
            with open(self.log_path, "w") as f:
                for item in self.error_list:
                    f.write("%s\n" % item)

            print(f"Error log exported to:  + {str(self.log_path)}")
            self.enable_buttons()


        def save_ags(self):
            self.disable_buttons()
            newFileName = filedialog.asksaveasfilename(filetypes=[('AGS Files', '*.ags')],defaultextension="*.ags",title="Save AGS file as...")
            if not newFileName:
                self.enable_buttons()
                return
            AGS4.dataframe_to_AGS4(self.tables, self.tables, newFileName)
            print('Done.')
            self.text.set('''AGS saved.
    ''')
            window.update()
            self.enable_buttons()

        def get_gint(self):
            self.disable_buttons()

            self.gint_location = filedialog.askopenfilename(filetypes=[('gINT Project', '*.gpj')], title="Please insert gINT database...")

            if not self.gint_location:
                messagebox.showwarning(title="Gimme a gINT!", message="You didn't select a gINT file.")
                self.enable_buttons()
                return

            self.text.set('''Getting gINT, please wait...
    ''')
            window.update()

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
            lab = self.lab_select.get()
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
                self.text.set('''Matching complete! Click: 'Save AGS file'.
    ''')
                window.update()
                print("Matching complete!")
                self.enable_buttons()
                if self.error == True:
                    self.text.set('''gINT matches, Lab doesn't.
    Re-open the AGS and select correct lab.''')
                    window.update()
            else:    
                self.text.set('''Couldn't match sample data.
    Did you select the correct gINT or AGS?''')
                window.update()
                print("Unable to match sample data from gINT.") 
                self.enable_buttons()


        def match_unique_id_gqm(self):
            self.disable_buttons()
            self.get_gint()
            self.matched = False
            self.error = False

            if not self.gint_location or self.gint_location == '':
                self.text.set('''AGS file loaded.
    ''')
                window.update()
                return

            self.text.set('''Matching AGS to gINT, please wait...
    ''')
            window.update()
            print(f"Matching GM Lab AGS to gINT... {self.gint_location}") 

            self.get_ags_tables()

            if 'GCHM' in self.ags_tables or 'ERES' in self.ags_tables:
                self.error = True
                print("GCHM or ERES table(s) found.")

            self.create_match_id()

            for table in self.ags_tables:
                try:
                    gint_rows = self.get_spec().shape[0]

                    for row in range (0,gint_rows):
                        self.get_spec()['match_id'][row] = str(self.get_spec()['PointID'][row]) + str(self.get_spec()['SPEC_REF'][row]) + str(format(self.get_spec()['Depth'][row],'.2f'))

                    for row in range (2,len(self.tables[table])):
                        self.tables[table]['match_id'][row] = str(self.tables[table]['LOCA_ID'][row]) + str(self.tables[table]['SAMP_TYPE'][row]) + str(self.tables[table]['SAMP_TOP'][row])
                        
                    if table == 'SPEC':
                        try:
                            for row in range (2,len(self.tables['SPEC'])):
                                self.tables['SPEC']['match_id'][row] = str(self.tables['SPEC']['LOCA_ID'][row]) + str(self.tables['SPEC']['SPEC_REF'][row]) + str(self.tables['SPEC']['SPEC_DPTH'][row])
                        except:
                            pass

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
                                        self.tables[table]['SPEC_DPTH'][tablerow] = format(self.get_spec()['Depth'][gintrow],'.2f')
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
                                    if self.tables['TREG']['TREG_TYPE'][tablerow] != 'CD':
                                        self.tables[table]['TRET_SHST'][tablerow] = round(float(self.tables[table]['TRET_DEVF'][tablerow]) / 2)
                            if 'TRET_CELL' in self.tables[table].keys():
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
                                        self.tables[table]['Depth'][tablerow] = format(self.get_spec()['Depth'][gintrow] + 0.01,'.2f')
                                    else:
                                        self.tables[table]['Depth'][tablerow] = format(self.get_spec()['Depth'][gintrow],'.2f')


                    '''RELD'''
                    if table == 'RELD':
                        if 'Depth' not in self.tables[table]:
                            self.tables[table].insert(8,'Depth','')
                        for tablerow in range(2,len(self.tables[table])):
                            for gintrow in range(0,gint_rows):
                                if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                    self.tables[table]['Depth'][tablerow] = format(self.get_spec()['Depth'][gintrow],'.2f')



                    '''RPLT'''
                    if table == 'RPLT':
                        if 'Depth' not in self.tables[table]:
                            self.tables[table].insert(8,'Depth','')
                        for tablerow in range(2,len(self.tables[table])):
                            for gintrow in range(0,gint_rows):
                                if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                    self.tables[table]['Depth'][tablerow] = format(self.get_spec()['Depth'][gintrow],'.2f')

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
                self.text.set('''AGS file loaded.
    ''')
                window.update()
                return

            self.text.set('''Matching AGS to gINT, please wait...
    ''')
            window.update()
            print(f"Matching DETS AGS to gINT... {self.gint_location}") 

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
                        self.get_spec()['match_id'][row] = str(self.get_spec()['PointID'][row]) + str(format(self.get_spec()['Depth'][row],'.2f'))

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

        
        def match_unique_id_soils(self):
            self.disable_buttons()
            self.get_gint()
            self.matched = False
            self.error = False

            if not self.gint_location or self.gint_location == '':
                self.text.set('''AGS file loaded.
    ''')
                window.update()
                return

            self.text.set('''Matching AGS to gINT, please wait...
    ''')
            window.update()
            print(f"Matching Structural Soils AGS to gINT... {self.gint_location}") 

            self.get_ags_tables()

            self.create_match_id()


            for table in self.ags_tables:
                try:
                    gint_rows = self.get_spec().shape[0]

                    for row in range (0,gint_rows):
                        self.get_spec()['match_id'][row] = str(self.get_spec()['PointID'][row]) + str(format(self.get_spec()['Depth'][row],'.2f'))

                    for row in range (2,len(self.tables[table])):
                        self.tables[table]['match_id'][row] = str(self.tables[table]['LOCA_ID'][row]) + str(self.tables[table]['SAMP_TOP'][row])

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
                                    self.tables[table]['SPEC_DPTH'][tablerow] = format(self.get_spec()['Depth'][gintrow],'.2f')
                                    
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
                self.text.set('''AGS file loaded.
    ''')
                window.update()
                return

            self.text.set('''Matching AGS to gINT, please wait...
    ''')
            window.update()
            print(f"Matching PSL AGS to gINT... {self.gint_location}") 

            self.get_ags_tables()

            self.create_match_id()

            for table in self.ags_tables:
                try:
                    gint_rows = self.get_spec().shape[0]

                    for row in range (0,gint_rows):
                        self.get_spec()['match_id'][row] = str(self.get_spec()['PointID'][row]) + str(format(self.get_spec()['Depth'][row],'.2f'))

                    for row in range (2,len(self.tables[table])):
                        self.tables[table]['match_id'][row] = str(self.tables[table]['LOCA_ID'][row]) + str(self.tables[table]['SAMP_TOP'][row])

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
                                    self.tables[table]['SPEC_DPTH'][tablerow] = format(self.get_spec()['Depth'][gintrow],'.2f')
                                    
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
                self.text.set('''AGS file loaded.
    ''')
                window.update()
                return

            self.text.set('''Matching AGS to gINT, please wait...
    ''')
            window.update()
            print(f"Matching Geolabs AGS to gINT... {self.gint_location}") 

            self.get_ags_tables()

            self.create_match_id()

            for table in self.ags_tables:
                try:
                    gint_rows = self.get_spec().shape[0]

                    for row in range (0,gint_rows):
                        self.get_spec()['match_id'][row] = str(self.get_spec()['PointID'][row]) + str(format(self.get_spec()['Depth'][row],'.2f'))

                    for row in range (2,len(self.tables[table])):
                        self.tables[table]['match_id'][row] = str(self.tables[table]['LOCA_ID'][row]) + str(self.tables[table]['SAMP_TOP'][row])

            
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
                                    self.tables[table]['SPEC_DPTH'][tablerow] = format(self.get_spec()['Depth'][gintrow],'.2f')
                                    
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
                self.text.set('''AGS file loaded.
    ''')
                window.update()
                return

            self.text.set('''Matching AGS to gINT, please wait...
    ''')
            window.update()
            print(f"Matching Geolabs AGS to gINT... {self.gint_location}") 

            self.get_ags_tables()

            self.create_match_id()

            for table in self.ags_tables:
                try:                
                    if 'Depth' not in self.tables[table]:
                        self.tables[table].insert(8,'Depth','')

                    gint_rows = self.get_spec().shape[0]

                    '''Using for Fugro Boreholes (50HZ samples have different SAMP format including dupe depths)'''
                    for row in range (0,gint_rows):
                        self.get_spec()['match_id'][row] = str(self.get_spec()['PointID'][row]) + str(format(self.get_spec()['SAMP_Depth'][row],'.2f')) + str(self.get_spec()['SAMP_REF'][row])

                    for row in range (2,len(self.tables[table])):
                        self.tables[table]['match_id'][row] = str(self.tables[table]['LOCA_ID'][row]) + str(self.tables[table]['SAMP_TOP'][row]) + str(self.tables[table]['SAMP_REF'][row])

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
                                    self.tables[table]['SAMP_TOP'][tablerow] = format(self.get_spec()['SAMP_Depth'][gintrow],'.2f')
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
                self.text.set('''AGS file loaded.
    ''')
                window.update()
                return

            self.text.set('''Matching AGS to gINT, please wait...
    ''')
            window.update()
            print(f"Matching GM Lab AGS to gINT... {self.gint_location}") 

            self.get_ags_tables()

            if 'GCHM' in self.ags_tables or 'ERES' in self.ags_tables:
                self.error = True
                print("GCHM or ERES table(s) found.")

            self.create_match_id()

            for table in self.ags_tables:
                try:
                    gint_rows = self.get_spec().shape[0]

                    for row in range (0,gint_rows):
                        self.get_spec()['match_id'][row] = str(self.get_spec()['PointID'][row]) + str(self.get_spec()['SPEC_REF'][row]) + str(format(self.get_spec()['Depth'][row],'.2f')) + str(self.get_spec()['SAMP_TYPE'][row][0])

                    for row in range (2,len(self.tables[table])):
                        self.tables[table]['match_id'][row] = str(self.tables[table]['LOCA_ID'][row]) + str(self.tables[table]['SAMP_TYPE'][row]) + str(self.tables[table]['SAMP_TOP'][row]) + str(self.tables[table]['SAMP_REF'][row][0])
                        
                    if table == 'SPEC':
                        try:
                            for row in range (2,len(self.tables['SPEC'])):
                                self.tables['SPEC']['match_id'][row] = str(self.tables['SPEC']['LOCA_ID'][row]) + str(self.tables['SPEC']['SPEC_REF'][row]) + str(self.tables['SPEC']['SPEC_DPTH'][row])
                        except:
                            pass

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
                                        self.tables[table]['SPEC_DPTH'][tablerow] = format(self.get_spec()['Depth'][gintrow],'.2f')
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
                                    if self.tables['TREG']['TREG_TYPE'][tablerow] != 'CD':
                                        self.tables[table]['TRET_SHST'][tablerow] = round(float(self.tables[table]['TRET_DEVF'][tablerow]) / 2)
                            if 'TRET_CELL' in self.tables[table].keys():
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
                                        self.tables[table]['Depth'][tablerow] = format(self.get_spec()['Depth'][gintrow] + 0.01,'.2f')
                                    else:
                                        self.tables[table]['Depth'][tablerow] = format(self.get_spec()['Depth'][gintrow],'.2f')


                    '''RELD'''
                    if table == 'RELD':
                        if 'Depth' not in self.tables[table]:
                            self.tables[table].insert(8,'Depth','')
                        for tablerow in range(2,len(self.tables[table])):
                            for gintrow in range(0,gint_rows):
                                if self.tables[table]['match_id'][tablerow] == self.get_spec()['match_id'][gintrow]:
                                    self.tables[table]['Depth'][tablerow] = format(self.get_spec()['Depth'][gintrow],'.2f')


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
                self.text.set('''AGS file loaded.
    ''')
                window.update()
                return

            self.text.set('''Matching AGS to gINT, please wait...
    ''')
            window.update()
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

                self.text.set('''CPT Only export ready.
    Click "Save AGS file"''')
                window.update()
                print("CPT Data export ready. Click 'Save AGS file'.")

            else:
                self.text.set('''Could not find any CPT tables.
    Check the AGS with "View data".''')
                window.update()
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

                self.text.set('''Lab Data & GEOL export ready.
    Click "Save AGS file"''')
                window.update()
                print("Lab Data & GEOL export ready. Click 'Save AGS file'.")
                self.ags_table_reset()

            else:
                self.text.set('''Could not find any Lab or GEOL tables.
    Check the AGS with "View data".''')
                window.update()
                print("No Lab or GEOL groups found - did this AGS contain CPT data? Check the data with 'View data'.")
                self.ags_table_reset()

        def disable_buttons(self):
            self.button_open.configure(state=tk.DISABLED)
            self.pandas_gui.configure(state=tk.DISABLED)
            self.button_count_results.configure(state=tk.DISABLED)
            self.button_ags_checker.configure(state=tk.DISABLED)
            self.button_save_ags.configure(state=tk.DISABLED)
            self.button_del_tbl.configure(state=tk.DISABLED)
            self.button_cpt_only.configure(state=tk.DISABLED)
            self.button_lab_only.configure(state=tk.DISABLED)
            self.lab_select.configure(state=tk.DISABLED)
            self.button_match_lab.configure(state=tk.DISABLED)
            
        def enable_buttons(self):
            self.button_open.configure(state=tk.NORMAL)
            self.pandas_gui.configure(state=tk.NORMAL)
            self.button_count_results.configure(state=tk.NORMAL)
            self.button_ags_checker.configure(state=tk.NORMAL)
            self.button_save_ags.configure(state=tk.NORMAL)
            self.button_del_tbl.configure(state=tk.NORMAL)
            self.button_cpt_only.configure(state=tk.NORMAL)
            self.button_lab_only.configure(state=tk.NORMAL)
            self.lab_select.configure(state=tk.NORMAL)
            self.button_match_lab.configure(state=tk.NORMAL)


# if not in any table flag warning for incorrect AGS instead of ERES/GCHM check for DETS
#SPLASH WINDOW NOT SCALING CORRECTLY ON 100% MAGNIFICATION - RESIZE WINDOW?

    window = ct.CTk()
    app = Application()
    window.mainloop()

if __name__ == '__main__':
    main()
