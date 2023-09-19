import pandas as pd
import numpy as np
from PyQt5.QtWidgets import QWidget
from PyQt5.QtCore import pyqtSignal
from statistics import mean

class LabHandler(QWidget):
    _update_text = pyqtSignal(str)
    _nice = pyqtSignal()
    _disable = pyqtSignal()
    _enable = pyqtSignal()

    def __init__(self):
        super(QWidget, self).__init__()
        self.tables: dict = None
        self.spec: pd.DataFrame = None
        self.ags_tables: list = []
        self.matched: bool = None

    def match_unique_id_gqm(self):
        self.matched = False
        self.error = False

        self.spec['Depth'] = self.spec['Depth'].map('{:,.2f}'.format)
        self.spec['Depth'] = self.spec['Depth'].astype(str)
        self.spec['match_id'] = self.spec['PointID']
        self.spec['match_id'] += self.spec['SPEC_REF']
        self.spec['match_id'] += self.spec['Depth']

        for table in self.ags_tables:
            try:
                gint_rows = self.spec.shape[0]

                self.tables[table]['match_id'] = self.tables[table]['LOCA_ID']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TYPE']
                self.tables[table]['match_id'] += self.tables[table]['SAMP_TOP']

                try:
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
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
                                self.tables[table]['SAMP_TOP'][tablerow] = format(self.spec['SAMP_Depth'][gintrow],'.2f')

                                try:
                                    self.tables[table]['SPEC_REF'][tablerow] = self.spec['SPEC_REF'][gintrow]
                                except:
                                    pass

                                try:
                                    self.tables[table]['SPEC_DPTH'][tablerow] = self.spec['Depth'][gintrow]
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
                        for gintrow in range(0,gint_rows):
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
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                                self.tables[table]['Depth'][tablerow] = self.spec['Depth'][gintrow]



                '''RPLT'''
                if table == 'RPLT':
                    if 'Depth' not in self.tables[table]:
                        self.tables[table].insert(8,'Depth','')
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                                self.tables[table]['Depth'][tablerow] = self.spec['Depth'][gintrow]
                            if "RPLT_FAIL" in self.tables[table]:
                                if "." in str(self.tables[table]['RPLT_FAIL'][tablerow]):
                                    self.tables[table]['RPLT_FAIL'][tablerow] = float(self.tables[table]['RPLT_FAIL'][tablerow] * 1000) 


                '''RDEN'''
                if table == 'RDEN':
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
                            if self.tables[table]['match_id'][tablerow] == self.spec['match_id'][gintrow]:
                                if float(self.tables[table]['RDEN_DDEN'][tablerow]) <= 0:
                                    self.tables[table]['RDEN_DDEN'][tablerow] = 0
                                    self.tables[table]['RDEN_PORO'][tablerow] = 0


                '''LDYN'''
                if table == 'LDYN':
                    for tablerow in range(2,len(self.tables[table])):
                        for gintrow in range(0,gint_rows):
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
                print(f"Couldn't find table or field, skipping... {str(e)}")
                pass

        self.check_matched_to_gint()




    def check_matched_to_gint(self):
        if self.matched:
            self._update_text.emit('''Matching complete! Check the data with 'View Data'
Click: 'Save AGS file'.''')
            print("Matching complete!")
            self._nice.emit()
            self._enable.emit()
            if self.error == True:
                self._update_text.emit('''gINT matches, Lab doesn't.
Re-open the AGS and select correct lab.''')
        else:    
            self._update_text.emit('''Couldn't match sample data.
Did you select the correct gINT or AGS?''')
            print("Unable to match sample data from gINT.") 
            self._enable.emit()