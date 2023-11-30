import pandas as pd
import numpy as np
from PyQt5.QtWidgets import QWidget
from PyQt5.QtCore import pyqtSignal
from statistics import mean
from rich import print as rprint

class LabHandler(QWidget):
    _update_text = pyqtSignal(str)
    _nice = pyqtSignal()
    _disable = pyqtSignal()
    _enable = pyqtSignal()
    _progress_max = pyqtSignal(int)
    _progress_current = pyqtSignal(int)

    def __init__(self):
        super(LabHandler, self).__init__()
        self.tables: dict = None
        self.spec: pd.DataFrame = None
        self.ags_tables: list = []
        self.matched: bool = None
        self.error: bool = None

    def check_matched_to_gint(self):
        if self.matched:
            self._update_text.emit('''Matching complete! Check the data with 'View Data'
Click: 'Save AGS file'.''')
            rprint(f"[green][bold]Matching complete![/bold][green]")
            self._nice.emit()
            self._enable.emit()
            if self.error == True:
                self._update_text.emit('''gINT matches, Lab doesn't.
Re-open the AGS and select correct lab.''')
        else:    
            self._update_text.emit('''Couldn't match sample data.
Did you select the correct gINT or AGS?''')
            rprint(f"[red][bold]Unable to match sample data from gINT.[/bold][red]")
            self._enable.emit()

    def remove_match_id(self):
        for table in self.ags_tables:
            if "match_id" in self.tables[table]:
                self.tables[table].drop(['match_id'], axis=1, inplace=True)

    def filter_spec(self):
        #since i'm forcing myself to use a nested loop to match ids, might as well reduce the size of SPEC to speed things by filtering it on only the samples in the ags
        bhs = [list(self.tables[table]['match_id'])[2:] for table in self.ags_tables if 'match_id' in self.tables[table]]
        bhs = list(set([x for y in bhs for x in y]))
        self.spec = self.spec[self.spec['match_id'].isin(bhs)]
        self.spec.reset_index(drop=True, inplace=True)
        if self.spec.shape[0] == 0:
            return rprint('[red][bold]NO MATCH TO GINT![bold][red]')


    def match_unique_id_gqm(self):
        self.matched = False
        self.error = False
        progress = 0
        progress_total = (len(self.tables.keys()) - 2) * 100
        self._progress_max.emit(progress_total)
        self._progress_current.emit(progress)   

        try:
            self.spec['Depth'] = self.spec['Depth'].map('{:,.2f}'.format)
            self.spec['Depth'] = self.spec['Depth'].astype(str)
            self.spec['match_id'] = self.spec['PointID']
            self.spec['match_id'] += self.spec['SPEC_REF']
            self.spec['match_id'] += self.spec['Depth']

            for table in self.ags_tables:
                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TYPE']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']

            self.filter_spec()

            for table in self.ags_tables:
                rprint(f"[yellow]Matching [bold]{table}[/bold]...[yellow]")

                for tablerow in range(2,len(self.tables[table])):
                    for gintrow in range(0,self.spec.shape[0]):
                        if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                            self.matched = True

                            if table == 'CONG':
                                if self.tables[table]['SPEC_REF'][tablerow] == "OED" or self.tables[table]['SPEC_REF'][tablerow] == "OEDR" and self.tables[table]['CONG_TYPE'][tablerow] == '':
                                    self.tables[table]['CONG_TYPE'][tablerow] = self.tables[table]['SPEC_REF'][tablerow]

                            if table == 'SAMP':
                                self.tables[table]['SAMP_REM'][tablerow] = self.spec['SPEC_REF'][gintrow]

                            self.tables[table]['SAMP_ID'][tablerow] = self.spec['SAMP_ID'][gintrow]
                            self.tables[table]['SAMP_REF'][tablerow] = self.spec['SAMP_REF'][gintrow]
                            self.tables[table]['SAMP_TYPE'][tablerow] = self.spec['SAMP_TYPE'][gintrow]
                            self.tables[table]['SAMP_TOP'][tablerow] = format(float(self.spec['SAMP_Depth'][gintrow]),'.2f')
                            self.tables[table]['SPEC_REF'][tablerow] = self.spec['SPEC_REF'][gintrow]
                            self.tables[table]['SPEC_DPTH'][tablerow] = self.spec['Depth'][gintrow]


                            for x in self.tables[table].keys():
                                if "LAB" in x:
                                    self.tables[table][x][tablerow] = "GM Lab"
                try:
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
                            self.tables[table]['TRET_DEVF'][tablerow] = int(round(float(self.tables[table]['TRET_DEVF'][tablerow]), 0))

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
                            for gintrow in range(0,self.spec.shape[0]):
                                if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                                    if self.tables['TRIG']['TRIG_COND'][tablerow] == 'REMOULDED':
                                        self.tables[table]['Depth'][tablerow] = round(float(self.spec['Depth'][gintrow]) + 0.01,2)
                                    else:
                                        self.tables[table]['Depth'][tablerow] = self.spec['Depth'][gintrow]

                    '''RELD'''
                    if table == 'RELD':
                        if 'Depth' not in self.tables[table]:
                            self.tables[table].insert(8,'Depth','')
                        for tablerow in range(2,len(self.tables[table])):
                            for gintrow in range(0,self.spec.shape[0]):
                                if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                                    self.tables[table]['Depth'][tablerow] = self.spec['Depth'][gintrow]

                    '''RPLT'''
                    if table == 'RPLT':
                        if 'Depth' not in self.tables[table]:
                            self.tables[table].insert(8,'Depth','')
                        for tablerow in range(2,len(self.tables[table])):
                            for gintrow in range(0,self.spec.shape[0]):
                                if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                                    self.tables[table]['Depth'][tablerow] = self.spec['Depth'][gintrow]
                                if "RPLT_FAIL" in self.tables[table]:
                                    if "." in str(self.tables[table]['RPLT_FAIL'][tablerow]):
                                        self.tables[table]['RPLT_FAIL'][tablerow] = float(self.tables[table]['RPLT_FAIL'][tablerow] * 1000) 

                    '''RDEN'''
                    if table == 'RDEN':
                        for tablerow in range(2,len(self.tables[table])):
                            for gintrow in range(0,self.spec.shape[0]):
                                if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                                    if float(self.tables[table]['RDEN_DDEN'][tablerow]) <= 0:
                                        self.tables[table]['RDEN_DDEN'][tablerow] = 0
                                        self.tables[table]['RDEN_PORO'][tablerow] = 0

                    '''LDYN'''
                    if table == 'LDYN':
                        for tablerow in range(2,len(self.tables[table])):
                            for gintrow in range(0,self.spec.shape[0]):
                                if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
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

                    '''LRES'''
                    if table == 'LRES':
                        for tablerow in range(2,len(self.tables[table])):
                            if self.tables[table]['LRES_TEMP'][tablerow] != '':
                                self.tables[table]['LRES_TEMP'][tablerow] = int(round(float(self.tables[table]['LRES_TEMP'][tablerow]),0))
                
                except Exception as e:
                    rprint(f'[red][b]ERROR[b][/red] in [red]{table}[/red]: Error: {e}')

                progress += 100
                self._progress_current.emit(progress)  

        except Exception as e:
            rprint(f"[red]ERROR[/red] matching in [red]{table}[/red]... Please check the data. Error: [white]{str(e)}[/white]")
            pass

        self.check_matched_to_gint()


    def match_unique_id_dets(self):
        self.matched = False
        self.error = False
        progress = 0
        progress_total = (len(self.tables.keys()) - 2) * 100
        self._progress_max.emit(progress_total)
        self._progress_current.emit(progress)   

        if 'GCHM' in self.ags_tables or 'ERES' in self.ags_tables:
            pass
        else:
            self.error = True
            print("Cannot find GCHM or ERES - looks like this AGS is from GM Lab.")

        try:
            self.spec['Depth'] = self.spec['Depth'].map('{:,.2f}'.format)
            self.spec['Depth'] = self.spec['Depth'].astype(str)
            self.spec['match_id'] = self.spec['PointID']
            self.spec['match_id'] += self.spec['Depth']

            for table in self.ags_tables:
                self.tables[table]['LOCA_ID'] = self.tables[table]['LOCA_ID'].str.split(" ", n=1, expand=True)[0]
                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']

            self.filter_spec()

            for table in self.ags_tables:
                rprint(f"[yellow]Matching [bold]{table}[/bold]...[yellow]")

                for tablerow in range(2,len(self.tables[table])):
                    for gintrow in range(0,self.spec.shape[0]):
                        if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                            self.matched = True
                            if table == 'ERES':
                                if 'ERES_REM' not in self.tables[table].keys():
                                    self.tables[table].insert(len(self.tables[table].keys()),'ERES_REM','')
                                self.tables[table]['ERES_REM'][tablerow] = self.tables[table]['SPEC_REF'][tablerow]
                            self.tables[table]['LOCA_ID'][tablerow] = self.spec['PointID'][gintrow]
                            self.tables[table]['SAMP_ID'][tablerow] = self.spec['SAMP_ID'][gintrow]
                            self.tables[table]['SAMP_REF'][tablerow] = self.spec['SAMP_REF'][gintrow]
                            self.tables[table]['SAMP_TYPE'][tablerow] = self.spec['SAMP_TYPE'][gintrow]
                            self.tables[table]['SPEC_REF'][tablerow] = self.spec['SPEC_REF'][gintrow]
                            self.tables[table]['SAMP_TOP'][tablerow] = format(float(self.spec['SAMP_Depth'][gintrow]),'.2f')
                            self.tables[table]['SPEC_DPTH'][tablerow] = self.spec['Depth'][gintrow]
                            
                            for x in self.tables[table].keys():
                                if "LAB" in x:
                                    self.tables[table][x][tablerow] = "DETS"
                try:
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
                    rprint(f'[red][b]ERROR[b][/red] in [red]{table}[/red]: Error: {e}')

                progress += 100
                self._progress_current.emit(progress)  

        except Exception as e:
            rprint(f"[red]ERROR[/red] matching in [red]{table}[/red]... Please check the data. Error: [white]{str(e)}[/white]")
            pass

        self.check_matched_to_gint()


    def match_unique_id_soils(self):
        self.matched = False
        self.error = False
        progress = 0
        progress_total = (len(self.tables.keys()) - 2) * 100
        self._progress_max.emit(progress_total)
        self._progress_current.emit(progress)   

        try:
            self.spec['Depth'] = self.spec['Depth'].map('{:,.2f}'.format)
            self.spec['Depth'] = self.spec['Depth'].astype(str)
            self.spec['match_id'] = self.spec['PointID']
            self.spec['match_id'] += self.spec['Depth']
            self.spec['SAMP_Depth'] = self.spec['SAMP_Depth'].astype(str)

            for table in self.ags_tables:
                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']

            self.filter_spec()

            for table in self.ags_tables:
                rprint(f"[yellow]Matching [bold]{table}[/bold]...[yellow]")

                for tablerow in range(2,len(self.tables[table])):
                    for gintrow in range(0,self.spec.shape[0]):
                        if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                            self.matched = True
                            self.tables[table]['LOCA_ID'][tablerow] = self.spec['PointID'][gintrow]
                            self.tables[table]['SAMP_ID'][tablerow] = self.spec['SAMP_ID'][gintrow]
                            self.tables[table]['SAMP_REF'][tablerow] = self.spec['SAMP_REF'][gintrow]
                            self.tables[table]['SAMP_TYPE'][tablerow] = self.spec['SAMP_TYPE'][gintrow]
                            self.tables[table]['SPEC_REF'][tablerow] = self.spec['SPEC_REF'][gintrow]
                            self.tables[table]['SAMP_TOP'][tablerow] = format(float(self.spec['SAMP_Depth'][gintrow]),'.2f')
                            self.tables[table]['SPEC_DPTH'][tablerow] = self.spec['Depth'][gintrow]
                            
                            for x in self.tables[table].keys():
                                if "LAB" in x:
                                    self.tables[table][x][tablerow] = "Structural Soils"
                try:                                    
                    '''CONG'''
                    if table == 'CONG':
                        for tablerow in range(2,len(self.tables[table])):
                            if "undisturbed" in str(self.tables[table]['CONG_COND'][tablerow].lower()):
                                self.tables[table]['CONG_COND'][tablerow] = "UNDISTURBED"
                            if "oed" in str(self.tables[table]['CONG_TYPE'][tablerow].lower()):
                                self.tables[table]['CONG_TYPE'][tablerow] = "IL OEDOMETER"
                                self.tables[table]['CONG_COND'][tablerow] = "UNDISTURBED"
                            if "#" in str(self.tables[table]['CONG_PDEN'][tablerow].lower()):
                                self.tables[table]['CONG_PDEN'][tablerow] = str(self.tables[table]['CONG_PDEN'][tablerow]).split('#')[1]
                
                except Exception as e:
                    rprint(f'[red][b]ERROR[b][/red] in [red]{table}[/red]: Error: {e}')
                
                progress += 100
                self._progress_current.emit(progress)  

        except Exception as e:
            rprint(f"[red]ERROR[/red] matching in [red]{table}[/red]... Please check the data. Error: [white]{str(e)}[/white]")
            pass

        self.check_matched_to_gint()


    def match_unique_id_psl(self):
        self.matched = False
        self.error = False
        progress = 0
        progress_total = (len(self.tables.keys()) - 2) * 100
        self._progress_max.emit(progress_total)
        self._progress_current.emit(progress)   

        try:
            self.spec['Depth'] = self.spec['Depth'].map('{:,.2f}'.format)
            self.spec['Depth'] = self.spec['Depth'].astype(str)
            self.spec['match_id'] = self.spec['PointID']
            self.spec['match_id'] += self.spec['Depth']

            for table in self.ags_tables:
                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']

            self.filter_spec()

            for table in self.ags_tables:
                rprint(f"[yellow]Matching [bold]{table}[/bold]...[yellow]")

                for tablerow in range(2,len(self.tables[table])):
                    for gintrow in range(0,self.spec.shape[0]):
                        if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                            self.matched = True
                            self.tables[table]['LOCA_ID'][tablerow] = self.spec['PointID'][gintrow]
                            self.tables[table]['SAMP_ID'][tablerow] = self.spec['SAMP_ID'][gintrow]
                            self.tables[table]['SAMP_REF'][tablerow] = self.spec['SAMP_REF'][gintrow]
                            self.tables[table]['SAMP_TYPE'][tablerow] = self.spec['SAMP_TYPE'][gintrow]
                            self.tables[table]['SPEC_REF'][tablerow] = self.spec['SPEC_REF'][gintrow]
                            self.tables[table]['SAMP_TOP'][tablerow] = format(float(self.spec['SAMP_Depth'][gintrow]),'.2f')
                            self.tables[table]['SPEC_DPTH'][tablerow] = self.spec['Depth'][gintrow]
                            
                            for x in self.tables[table].keys():
                                if "LAB" in x:
                                    self.tables[table][x][tablerow] = "PSL"
                try:
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
                            if self.tables[table]['TRET_SHST'][tablerow] == self.tables[table]['TRET_DEVF'][tablerow]:
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
                    rprint(f'[red][b]ERROR[b][/red] in [red]{table}[/red]: Error: {e}')

                progress += 100
                self._progress_current.emit(progress)  

        except Exception as e:
            rprint(f"[red]ERROR[/red] matching in [red]{table}[/red]... Please check the data. Error: [white]{str(e)}[/white]")
            pass

        self.check_matched_to_gint()


    def match_unique_id_geolabs(self):
        self.matched = False
        self.error = False
        progress = 0
        progress_total = (len(self.tables.keys()) - 2) * 100
        self._progress_max.emit(progress_total)
        self._progress_current.emit(progress)   

        try:
            self.spec['Depth'] = self.spec['Depth'].map('{:,.2f}'.format)
            self.spec['Depth'] = self.spec['Depth'].astype(str)
            self.spec['match_id'] = self.spec['PointID']
            self.spec['match_id'] += self.spec['Depth']

            for table in self.ags_tables:
                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']

            self.filter_spec()

            for table in self.ags_tables:
                rprint(f"[yellow]Matching [bold]{table}[/bold]...[yellow]")

                for tablerow in range(2,len(self.tables[table])):
                    for gintrow in range(0,self.spec.shape[0]):
                        if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                            self.matched = True
                            self.tables[table]['LOCA_ID'][tablerow] = self.spec['PointID'][gintrow]
                            self.tables[table]['SAMP_ID'][tablerow] = self.spec['SAMP_ID'][gintrow]
                            self.tables[table]['SAMP_REF'][tablerow] = self.spec['SAMP_REF'][gintrow]
                            self.tables[table]['SAMP_TYPE'][tablerow] = self.spec['SAMP_TYPE'][gintrow]
                            self.tables[table]['SPEC_REF'][tablerow] = self.spec['SPEC_REF'][gintrow]
                            self.tables[table]['SAMP_TOP'][tablerow] = format(float(self.spec['SAMP_Depth'][gintrow]),'.2f')
                            self.tables[table]['SPEC_DPTH'][tablerow] = self.spec['Depth'][gintrow]
                            
                            for x in self.tables[table].keys():
                                if "LAB" in x:
                                    self.tables[table][x][tablerow] = "Geolabs Limited"
                try:
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
                    rprint(f'[red][b]ERROR[b][/red] in [red]{table}[/red]: Error: {e}')

                progress += 100
                self._progress_current.emit(progress)  

        except Exception as e:
            rprint(f"[red]ERROR[/red] matching in [red]{table}[/red]... Please check the data. Error: [white]{str(e)}[/white]")
            pass

        self.check_matched_to_gint()


    def match_unique_id_geolabs_fugro(self):
        self.matched = False
        self.error = False
        progress = 0
        progress_total = (len(self.tables.keys()) - 2) * 100
        self._progress_max.emit(progress_total)
        self._progress_current.emit(progress)   
        
        try:
            '''Using for Fugro Boreholes (50HZ samples have different SAMP format including dupe depths)'''
            self.spec['SAMP_Depth'] = self.spec['SAMP_Depth'].map('{:,.2f}'.format)
            self.spec['SAMP_Depth'] = self.spec['SAMP_Depth'].astype(str)
            self.spec['match_id'] = self.spec['PointID']
            self.spec['match_id'] += self.spec['SAMP_Depth']
            self.spec['match_id'] += self.spec['SAMP_REF']

            for table in self.ags_tables:
                if 'Depth' not in self.tables[table]:
                    self.tables[table].insert(8,'Depth','')

                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_REF']

            self.filter_spec()

            for table in self.ags_tables:
                rprint(f"[yellow]Matching [bold]{table}[/bold]...[yellow]")

                for tablerow in range(2,len(self.tables[table])):
                    for gintrow in range(0,self.spec.shape[0]):
                        if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                            self.matched = True
                            self.tables[table]['Depth'][tablerow] = self.tables[table]['SPEC_DPTH'][tablerow]
                            self.tables[table]['LOCA_ID'][tablerow] = self.spec['PointID'][gintrow]
                            self.tables[table]['SAMP_ID'][tablerow] = self.spec['SAMP_ID'][gintrow]
                            self.tables[table]['SAMP_REF'][tablerow] = self.spec['SAMP_REF'][gintrow]
                            self.tables[table]['SAMP_TYPE'][tablerow] = self.spec['SAMP_TYPE'][gintrow]
                            self.tables[table]['SPEC_REF'][tablerow] = self.spec['SPEC_REF'][gintrow]
                            self.tables[table]['SAMP_TOP'][tablerow] = self.spec['SAMP_Depth'][gintrow]
                            self.tables[table]['SPEC_DPTH'][tablerow] = format(float(self.spec['Depth'][gintrow]),'.2f')   
                try:
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
                    rprint(f'[red][b]ERROR[b][/red] in [red]{table}[/red]: Error: {e}')

                progress += 100
                self._progress_current.emit(progress)  

        except Exception as e:
            rprint(f"[red]ERROR[/red] matching in [red]{table}[/red]... Please check the data. Error: [white]{str(e)}[/white]")
            pass

        self.check_matched_to_gint()


    def match_unique_id_soils_pez(self):
        self.matched = False
        self.error = False
        progress = 0
        progress_total = (len(self.tables.keys()) - 2) * 100
        self._progress_max.emit(progress_total)
        self._progress_current.emit(progress)   

        try:
            self.spec['Depth'] = self.spec['Depth'].map('{:,.2f}'.format)
            self.spec['Depth'] = self.spec['Depth'].astype(str)
            self.spec['match_id'] = self.spec['PointID']
            self.spec['match_id'] += self.spec['Depth']
            self.spec['batched'] = self.spec['SAMP_TYPE'].astype(str).str[0]
            self.spec['match_id'] += self.spec['batched']
            self.spec.drop(['batched'], axis=1, inplace=True)
            
            for table in self.ags_tables:
                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SPEC_DPTH']
                self.tables[table]['batched'] = self.tables[table]['SAMP_TYPE'].astype(str).str[0]
                self.tables[table]['match_id'] += self.tables[table]['batched']
                self.tables[table].drop(['batched'], axis=1, inplace=True)

            self.filter_spec()

            for table in self.ags_tables:
                rprint(f"[yellow]Matching [bold]{table}[/bold]...[yellow]")

                for tablerow in range(2,len(self.tables[table])):
                    for gintrow in range(0,self.spec.shape[0]):
                        if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                            self.matched = True
                            self.tables[table]['LOCA_ID'][tablerow] = self.spec['PointID'][gintrow]
                            self.tables[table]['SAMP_ID'][tablerow] = self.spec['SAMP_ID'][gintrow]
                            self.tables[table]['SAMP_REF'][tablerow] = self.spec['SAMP_REF'][gintrow]
                            self.tables[table]['SAMP_TYPE'][tablerow] = self.spec['SAMP_TYPE'][gintrow]
                            self.tables[table]['SPEC_REF'][tablerow] = self.spec['SPEC_REF'][gintrow]
                            self.tables[table]['SAMP_TOP'][tablerow] = format(float(self.spec['SAMP_Depth'][gintrow]),'.2f')
                            self.tables[table]['SPEC_DPTH'][tablerow] = self.spec['Depth'][gintrow]
                            
                            for x in self.tables[table].keys():
                                if "LAB" in x:
                                    self.tables[table][x][tablerow] = "Structural Soils Ltd - Bristol Geotech lab"
                try:
                    '''CONG'''
                    if table == 'CONG':
                        for tablerow in range(2,len(self.tables[table])):
                            if "undisturbed" in str(self.tables[table]['CONG_COND'][tablerow].lower()):
                                self.tables[table]['CONG_COND'][tablerow] = "UNDISTURBED"
                            if "oed" in str(self.tables[table]['CONG_TYPE'][tablerow].lower()):
                                self.tables[table]['CONG_TYPE'][tablerow] = "IL OEDOMETER"
                                self.tables[table]['CONG_COND'][tablerow] = "UNDISTURBED"
                            if "#" in str(self.tables[table]['CONG_PDEN'][tablerow].lower()):
                                self.tables[table]['CONG_PDEN'][tablerow] = str(self.tables[table]['CONG_PDEN'][tablerow]).split('#')[1]

                    '''IRSG'''
                    if table == 'IRSG':
                        for tablerow in range(2,len(self.tables[table])):
                            if 'IRSG_COND' in self.tables[table]:
                                self.tables[table]['IRSG_COND'][tablerow] = str(self.tables[table]['IRSG_COND'][tablerow]).upper()

                    '''LDYN'''
                    if table == 'LDYN':
                        for tablerow in range(2,len(self.tables[table])):
                            self.tables[table]['LDYN_SG'][tablerow] = int(float(self.tables[table]['LDYN_SG'][tablerow]))

                    '''SHBT'''
                    if table == 'SHBT':
                        for tablerow in range(2,len(self.tables[table])):
                            if not self.tables[table]['SHBT_PDIN'][tablerow] == "" and float(self.tables[table]['SHBT_PDIN'][tablerow]) < 0:
                                self.tables[table]['SHBT_PDIN'][tablerow] = 0
                            if "#" in str(self.tables[table]['SHBT_PDEN'][tablerow].lower()):
                                self.tables[table]['SHBT_PDEN'][tablerow] = str(self.tables[table]['SHBT_PDEN'][tablerow]).split('#')[1]
                
                except Exception as e:
                    rprint(f'[red][b]ERROR[b][/red] in [red]{table}[/red]: Error: {e}')
                        
                progress += 100
                self._progress_current.emit(progress)  

        except Exception as e:
            rprint(f"[red]ERROR[/red] matching in [red]{table}[/red]... Please check the data. Error: [white]{str(e)}[/white]")
            pass

        self.check_matched_to_gint()


    def match_unique_id_gqm_pez(self):
        self.matched = False
        self.error = False
        progress = 0
        progress_total = (len(self.tables.keys()) - 2) * 100
        self._progress_max.emit(progress_total)
        self._progress_current.emit(progress)   

        if 'GCHM' in self.ags_tables or 'ERES' in self.ags_tables:
            self.error = True
            print("GCHM or ERES table(s) found.")

        try:
            self.spec['Depth'] = self.spec['Depth'].map('{:,.2f}'.format)
            self.spec['Depth'] = self.spec['Depth'].astype(str)
            self.spec['match_id'] = self.spec['PointID']
            self.spec['match_id'] += self.spec['SPEC_REF']
            self.spec['match_id'] += self.spec['Depth']
            self.spec['batched'] = self.spec['SAMP_TYPE'].astype(str).str[0]
            self.spec['match_id'] += self.spec['batched']
            self.spec.drop(['batched'], axis=1, inplace=True)

            for table in self.ags_tables:
                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TYPE']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']
                self.tables[table]['batched'] = self.tables[table]['SAMP_REF'].astype(str).str[0]
                self.tables[table]['match_id'] += self.tables[table]['batched']
                self.tables[table].drop(['batched'], axis=1, inplace=True)

            self.filter_spec()

            for table in self.ags_tables:
                rprint(f"[yellow]Matching [bold]{table}[/bold]...[yellow]")

                for tablerow in range(2,len(self.tables[table])):
                    for gintrow in range(0,self.spec.shape[0]):
                        if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                            self.matched = True

                            if table == 'CONG':
                                if self.tables[table]['SPEC_REF'][tablerow] == "OED" or self.tables[table]['SPEC_REF'][tablerow] == "OEDR" and self.tables[table]['CONG_TYPE'][tablerow] == '':
                                    self.tables[table]['CONG_TYPE'][tablerow] = self.tables[table]['SPEC_REF'][tablerow]

                            if table == 'SAMP':
                                self.tables[table]['SAMP_REM'][tablerow] = self.spec['SPEC_REF'][gintrow]

                            self.tables[table]['SAMP_ID'][tablerow] = self.spec['SAMP_ID'][gintrow]
                            self.tables[table]['SAMP_REF'][tablerow] = self.spec['SAMP_REF'][gintrow]
                            self.tables[table]['SAMP_TYPE'][tablerow] = self.spec['SAMP_TYPE'][gintrow]
                            self.tables[table]['SAMP_TOP'][tablerow] = format(float(self.spec['SAMP_Depth'][gintrow]),'.2f')
                            self.tables[table]['SPEC_REF'][tablerow] = self.spec['SPEC_REF'][gintrow]
                            self.tables[table]['SPEC_DPTH'][tablerow] = self.spec['Depth'][gintrow]
                            for x in self.tables[table].keys():
                                if "LAB" in x:
                                    self.tables[table][x][tablerow] = "GM Lab"
                try:
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
                            for gintrow in range(0,self.spec.shape[0]):
                                if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                                    if self.tables['TRIG']['TRIG_COND'][tablerow] == 'REMOULDED':
                                        self.tables[table]['Depth'][tablerow] = round(float(self.spec['Depth'][gintrow]) + 0.01,2)
                                    else:
                                        self.tables[table]['Depth'][tablerow] = self.spec['Depth'][gintrow]

                    '''RELD'''
                    if table == 'RELD':
                        if 'Depth' not in self.tables[table]:
                            self.tables[table].insert(8,'Depth','')
                        for tablerow in range(2,len(self.tables[table])):
                            for gintrow in range(0,self.spec.shape[0]):
                                if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                                    self.tables[table]['Depth'][tablerow] = self.spec['Depth'][gintrow]

                    '''LDYN'''
                    if table == 'LDYN':
                        for tablerow in range(2,len(self.tables[table])):
                            for gintrow in range(0,self.spec.shape[0]):
                                if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
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
                    rprint(f'[red][b]ERROR[b][/red] in [red]{table}[/red]: Error: {e}')

                progress += 100
                self._progress_current.emit(progress)  

        except Exception as e:
            rprint(f"[red]ERROR[/red] matching in [red]{table}[/red]... Please check the data. Error: [white]{str(e)}[/white]")
            pass

        self.check_matched_to_gint()


    def match_unique_id_dets_pez(self):
        self.matched = False
        self.error = False
        progress = 0
        progress_total = (len(self.tables.keys()) - 2) * 100
        self._progress_max.emit(progress_total)
        self._progress_current.emit(progress)   

        if 'GCHM' in self.ags_tables or 'ERES' in self.ags_tables:
            pass
        else:
            self.error = True
            print("Cannot find GCHM or ERES - looks like this AGS is from GM Lab.")

        try:
            self.spec['Depth'] = self.spec['Depth'].map('{:,.2f}'.format)
            self.spec['Depth'] = self.spec['Depth'].astype(str)
            self.spec['match_id'] = self.spec['PointID']
            self.spec['match_id'] += self.spec['Depth']
            self.spec['batched'] = self.spec['SAMP_TYPE'].astype(str).str[0]
            self.spec['match_id'] += self.spec['batched']
            self.spec.drop(['batched'], axis=1, inplace=True)
            self.spec['match_id'] += self.spec['SPEC_REF']

            for table in self.ags_tables:
                self.tables[table]['LOCA_ID'] = self.tables[table]['LOCA_ID'].str.split(" ", n=1, expand=True)[0]
                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']
                self.tables[table]['batched'] = self.tables[table]['SAMP_REF'].astype(str).str[0]
                self.tables[table]['match_id'] += self.tables[table]['batched']
                self.tables[table].drop(['batched'], axis=1, inplace=True)
                self.tables[table]['SAMP_REF'] = self.tables[table]['SAMP_REF'].str.split(" ", n=1, expand=True)[1]
                self.tables[table]['match_id'] += self.tables[table]['SAMP_REF']

            self.filter_spec()

            for table in self.ags_tables:
                rprint(f"[yellow]Matching [bold]{table}[/bold]...[yellow]")

                for tablerow in range(2,len(self.tables[table])):
                    for gintrow in range(0,self.spec.shape[0]):
                        if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                            self.matched = True
                            if table == 'ERES':
                                if 'ERES_REM' not in self.tables[table].keys():
                                    self.tables[table].insert(len(self.tables[table].keys()),'ERES_REM','')
                                self.tables[table]['ERES_REM'][tablerow] = self.tables[table]['SPEC_REF'][tablerow]
                            self.tables[table]['LOCA_ID'][tablerow] = self.spec['PointID'][gintrow]
                            self.tables[table]['SAMP_ID'][tablerow] = self.spec['SAMP_ID'][gintrow]
                            self.tables[table]['SAMP_REF'][tablerow] = self.spec['SAMP_REF'][gintrow]
                            self.tables[table]['SAMP_TYPE'][tablerow] = self.spec['SAMP_TYPE'][gintrow]
                            self.tables[table]['SPEC_REF'][tablerow] = self.spec['SPEC_REF'][gintrow]
                            self.tables[table]['SAMP_TOP'][tablerow] = format(float(self.spec['SAMP_Depth'][gintrow]),'.2f')
                            self.tables[table]['SPEC_DPTH'][tablerow] = self.spec['Depth'][gintrow]
                            
                            for x in self.tables[table].keys():
                                if "LAB" in x:
                                    self.tables[table][x][tablerow] = "DETS"
                try:
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
                    rprint(f'[red][b]ERROR[b][/red] in [red]{table}[/red]: Error: {e}')

                progress += 100
                self._progress_current.emit(progress)  

        except Exception as e:
            rprint(f"[red]ERROR[/red] matching in [red]{table}[/red]... Please check the data. Error: [white]{str(e)}[/white]")
            pass

        self.check_matched_to_gint()


    def match_unique_id_sinotech(self):
        self.matched = False
        self.error = False
        progress = 0
        progress_total = (len(self.tables.keys()) - 2) * 100
        self._progress_max.emit(progress_total)
        self._progress_current.emit(progress)

        try:
            self.spec['Depth'] = self.spec['Depth'].map('{:,.2f}'.format)
            self.spec['Depth'] = self.spec['Depth'].astype(str)
            self.spec['match_id'] = self.spec['PointID']
            self.spec['match_id'] += self.spec['Depth']

            for table in self.ags_tables:
                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']

            self.filter_spec()

            for table in self.ags_tables:
                rprint(f"[yellow]Matching [bold]{table}[/bold]...[yellow]")

                for tablerow in range(2,len(self.tables[table])):
                    for gintrow in range(0,self.spec.shape[0]):
                        if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                            self.matched = True
                            self.tables[table]['LOCA_ID'][tablerow] = self.spec['PointID'][gintrow]
                            self.tables[table]['SAMP_ID'][tablerow] = self.spec['SAMP_ID'][gintrow]
                            self.tables[table]['SAMP_REF'][tablerow] = self.spec['SAMP_REF'][gintrow]
                            self.tables[table]['SAMP_TYPE'][tablerow] = self.spec['SAMP_TYPE'][gintrow]
                            self.tables[table]['SPEC_REF'][tablerow] = self.spec['SPEC_REF'][gintrow]
                            self.tables[table]['SAMP_TOP'][tablerow] = format(float(self.spec['SAMP_Depth'][gintrow]),'.2f')
                            self.tables[table]['SPEC_DPTH'][tablerow] = self.spec['Depth'][gintrow]
                            
                            for x in self.tables[table].keys():
                                if "LAB" in x:
                                    self.tables[table][x][tablerow] = "Sinotech"
                try:
                    '''CONG'''
                    if table == 'CONG':
                        if 'CONG_TYPE' not in self.tables[table]:
                            self.tables[table].insert(10,'CONG_TYPE','')
                        for tablerow in range(2,len(self.tables[table])):
                            if "crs" in str(self.tables[table]['FILE_FSET'][tablerow].lower()):
                                self.tables[table]['CONG_COND'][tablerow] = "UNDISTURBED"
                                self.tables[table]['CONG_TYPE'][tablerow] = "CRS"
                            if "oed" in str(self.tables[table]['FILE_FSET'][tablerow].lower()):
                                self.tables[table]['CONG_TYPE'][tablerow] = "IL OEDOMETER"
                                
                    '''LLPL'''
                    if table == 'LLPL':
                        if 'Non-Plastic' not in self.tables[table]:
                            self.tables[table].insert(13,'Non-Plastic','')
                        for tablerow in range(2,len(self.tables[table])):
                            if self.tables[table]['LLPL_LL'][tablerow] == '' and self.tables[table]['LLPL_PL'][tablerow] == '' and self.tables[table]['LLPL_PI'][tablerow] == '' or self.tables[table]['LLPL_LL'][tablerow] == "NP":
                                self.tables[table]['Non-Plastic'][tablerow] = -1
                                
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
                            for gintrow in range(0,self.spec.shape[0]):
                                if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                                    if self.tables['TRIG']['TRIG_COND'][tablerow] == 'REMOULDED':
                                        self.tables[table]['Depth'][tablerow] = round(float(self.spec['Depth'][gintrow]) + 0.01,2)
                                    else:
                                        self.tables[table]['Depth'][tablerow] = self.spec['Depth'][gintrow]
                                        
                    '''TRET'''
                    if table == 'TRET':
                        for tablerow in range(2,len(self.tables[table])):
                            if float(self.tables[table]['TRET_DDEN'][tablerow]) > 4.0:
                                self.tables[table]['TRET_DDEN'][tablerow] = round(float(self.tables[table]['TRET_DDEN'][tablerow]) / 9.81, 2)
                                
                    '''RELD'''
                    if table == 'RELD':
                        for tablerow in range(2,len(self.tables[table])):
                            if float(self.tables[table]['RELD_DMAX'][tablerow]) > 4.0:
                                self.tables[table]['RELD_DMAX'][tablerow] = float(self.tables[table]['RELD_DMAX'][tablerow]) / 900.81
                                self.tables[table]['RELD_DMIN'][tablerow] = float(self.tables[table]['RELD_DMIN'][tablerow]) / 900.81
                                
                    '''LDEN'''
                    if table == 'LDEN':
                        for tablerow in range(2,len(self.tables[table])):
                            if not self.tables[table]['LDEN_BDEN'][tablerow] == "":
                                if float(self.tables[table]['LDEN_BDEN'][tablerow]) > 4.0:
                                    self.tables[table]['LDEN_BDEN'][tablerow] = float(self.tables[table]['LDEN_BDEN'][tablerow]) / 9.81
                            if not self.tables[table]['LDEN_DDEN'][tablerow] == "":
                                if float(self.tables[table]['LDEN_DDEN'][tablerow]) > 4.0:
                                    self.tables[table]['LDEN_DDEN'][tablerow] = float(self.tables[table]['LDEN_DDEN'][tablerow]) / 9.81

                except Exception as e:
                    rprint(f'[red][b]ERROR[b][/red] in [red]{table}[/red]: Error: {e}')

                progress += 100
                self._progress_current.emit(progress)  

        except Exception as e:
            rprint(f"[red]ERROR[/red] matching in [red]{table}[/red]... Please check the data. Error: [white]{str(e)}[/white]")
            pass

        self.check_matched_to_gint()

    def match_unique_id_mewo(self):
        self.matched = False
        self.error = False
        progress = 0
        progress_total = (len(self.tables.keys()) - 2) * 100
        self._progress_max.emit(progress_total)
        self._progress_current.emit(progress)

        try:
            self.spec['SPEC_DEPTH2'] = self.spec['SPEC_DEPTH2'].map('{:,.2f}'.format)
            self.spec['SPEC_DEPTH2'] = self.spec['SPEC_DEPTH2'].astype(str)
            self.spec['match_id'] = self.spec['PointID']
            self.spec['match_id'] += self.spec['SPEC_DEPTH2']

            for table in self.ags_tables:
                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SPEC_DPTH']

            self.filter_spec()

            for table in self.ags_tables:
                rprint(f"[yellow]Matching [bold]{table}[/bold]...[yellow]")

                for tablerow in range(2,len(self.tables[table])):
                    for gintrow in range(0,self.spec.shape[0]):
                        if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                            self.matched = True
                            self.tables[table]['LOCA_ID'][tablerow] = self.spec['PointID'][gintrow]
                            self.tables[table]['SAMP_ID'][tablerow] = self.spec['SAMP_ID'][gintrow]
                            self.tables[table]['SAMP_REF'][tablerow] = self.spec['SAMP_REF'][gintrow]
                            self.tables[table]['SAMP_TYPE'][tablerow] = self.spec['SAMP_TYPE'][gintrow]
                            #self.tables[table]['SPEC_REF'][tablerow] = self.spec['SPEC_REF'][gintrow]
                            self.tables[table]['SAMP_TOP'][tablerow] = format(float(self.spec['SAMP_Depth'][gintrow]),'.2f')
                            self.tables[table]['SPEC_DPTH'][tablerow] = self.spec['SPEC_DEPTH2'][gintrow]
                            
                            for x in self.tables[table].keys():
                                if "LAB" in x:
                                    self.tables[table][x][tablerow] = "Mewo"
                try:
                    '''TXTG'''
                    if table == 'TXTG':
                        for tablerow in range(2,len(self.tables[table])):
                            if "cd" in str(self.tables[table]['TXTG_TYPE'][tablerow].lower()):
                                self.tables[table]['TXTG_TYPE'][tablerow] = "CID"
                            if "cuc" in str(self.tables[table]['TXTG_TYPE'][tablerow].lower()):
                                self.tables[table]['TXTG_TYPE'][tablerow] = "CAUc"
                            if "cue" in str(self.tables[table]['TXTG_TYPE'][tablerow].lower()):
                                self.tables[table]['TXTG_TYPE'][tablerow] = "CAUe"
                
                except Exception as e:
                    rprint(f'[red][b]ERROR[b][/red] in [red]{table}[/red]: Error: {e}')

                progress += 100
                self._progress_current.emit(progress)    

        except Exception as e:
            rprint(f"[red]ERROR[/red] matching in [red]{table}[/red]... Please check the data. Error: [white]{str(e)}[/white]")
            pass

        self.check_matched_to_gint()

    def match_unique_id_Enviro(self):
        self.matched = False
        self.error = False
        progress = 0
        progress_total = (len(self.tables.keys()) - 2) * 100
        self._progress_max.emit(progress_total)
        self._progress_current.emit(progress)   

        if 'GCHM' in self.ags_tables or 'ERES' in self.ags_tables:
            pass
        else:
            self.error = True
            print("Cannot find GCHM or ERES - looks like this AGS is from GM Lab.")

        try:
            self.spec['Depth'] = self.spec['Depth'].map('{:,.2f}'.format)
            self.spec['Depth'] = self.spec['Depth'].astype(str)
            self.spec['match_id'] = self.spec['PointID']
            self.spec['match_id'] += self.spec['Depth']

            for table in self.ags_tables:
                self.tables[table]['LOCA_ID'] = self.tables[table]['LOCA_ID'].str.split("-", n=1, expand=True)[0]
                
                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SPEC_DPTH']

            self.filter_spec()

            for table in self.ags_tables:
                rprint(f"[yellow]Matching [bold]{table}[/bold]...[yellow]")

                for tablerow in range(2,len(self.tables[table])):
                    for gintrow in range(0,self.spec.shape[0]):
                        if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                            self.matched = True
                            if table == 'ERES':
                                if 'ERES_REM' not in self.tables[table].keys():
                                    self.tables[table].insert(len(self.tables[table].keys()),'ERES_REM','')
                                self.tables[table]['ERES_REM'][tablerow] = self.tables[table]['SPEC_REF'][tablerow]
                            self.tables[table]['LOCA_ID'][tablerow] = self.spec['PointID'][gintrow]
                            self.tables[table]['SAMP_ID'][tablerow] = self.spec['SAMP_ID'][gintrow]
                            self.tables[table]['SAMP_REF'][tablerow] = self.spec['SAMP_REF'][gintrow]
                            self.tables[table]['SAMP_TYPE'][tablerow] = self.spec['SAMP_TYPE'][gintrow]
                            self.tables[table]['SPEC_REF'][tablerow] = self.spec['SPEC_REF'][gintrow]
                            self.tables[table]['SAMP_TOP'][tablerow] = format(float(self.spec['SAMP_Depth'][gintrow]),'.2f')
                            self.tables[table]['SPEC_DPTH'][tablerow] = self.spec['Depth'][gintrow]
                            
                            for x in self.tables[table].keys():
                                if "LAB" in x:
                                    self.tables[table][x][tablerow] = "Enviro"
                try:
                    '''GCHM'''
                    if table == 'GCHM':
                        for tablerow in range(2,len(self.tables[table])):
                            if "ph" in str(self.tables[table]['GCHM_UNIT'][tablerow].lower()):
                                self.tables[table]['GCHM_UNIT'][tablerow] = "-"
                            if "co3" in str(self.tables[table]['GCHM_CODE'][tablerow].lower()):
                                self.tables[table]['GCHM_CODE'][tablerow] = "CACO3"

                    '''ERES'''
                    if table == 'ERES':
                        self.tables[table]['ERES_TNAM'] = self.tables[table]['ERES_NAME']
                        for tablerow in range(2,len(self.tables[table])):
                            if "<" in str(self.tables[table]['ERES_RTXT'][tablerow].lower()):
                                self.tables[table]['ERES_RTXT'][tablerow] = str(self.tables[table]['ERES_RTXT'][tablerow]).rsplit("<", 1)[1]
                            if "solid" in str(self.tables[table]['ERES_MATX'][tablerow].lower()):
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
                            if "stones" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                                self.tables[table]['ERES_TNAM'][tablerow] = "% Stones"
                            if "chloride" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                                self.tables[table]['ERES_TNAM'][tablerow] = "Cl"
                            if "los" in str(self.tables[table]['ERES_TNAM'][tablerow].lower()):
                                self.tables[table]['ERES_TNAM'][tablerow] = "LOI"
                            if "ph" in str(self.tables[table]['ERES_RUNI'][tablerow].lower()):
                                self.tables[table]['ERES_RUNI'][tablerow] = "-"
                            if "%" in str(self.tables[table]['ERES_RUNI'][tablerow].lower()):
                                self.tables[table]['ERES_RUNI'][tablerow] = "%"

                except Exception as e:
                    rprint(f'[red][b]ERROR[b][/red] in [red]{table}[/red]: Error: {e}')

                progress += 100
                self._progress_current.emit(progress)  

        except Exception as e:
            rprint(f"[red]ERROR[/red] matching in [red]{table}[/red]... Please check the data. Error: [white]{str(e)}[/white]")
            pass

        self.check_matched_to_gint()
