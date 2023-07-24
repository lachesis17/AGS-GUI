import pandas as pd
from PyQt5.QtCore import *
from PyQt5.QtWidgets import QApplication, QTableView
from PyQt5.QtGui import QKeySequence, QMouseEvent
import csv
import io
from typing_extensions import Literal
import sys
sys.stdout.reconfigure(encoding='utf-8')

class PandasModel(QAbstractTableModel):
    def __init__(self, dataframe: pd.DataFrame):
        super().__init__()
        '''saving some commands to be used on subclass'''
        #self.installEventFilter(self)
        #self.table.setSortingEnabled(True)
        #self.table.horizontalHeader().sectionPressed.connect(self.table.selectColumn)
        
        self.df = dataframe
        self.sort_state = 'None'
        
    def rowCount(self, parent: QPersistentModelIndex) -> int:
        if self.df is None:
            return 0
        else:
            return self.df.shape[0]

    def columnCount(self, parent=QModelIndex) -> int:
        if self.df is None:
            return 0
        else:
            return self.df.shape[1]

    def data(self, index, role: int):
        if index.isValid():
            if role == Qt.DisplayRole or role == Qt.EditRole:
                return str(self.df.iloc[index.row(), index.column()])
        return None

    def setData(self, index, value, role):
        try:
            value = float(value)
        except ValueError:
            pass
        
        if role == Qt.EditRole:
            self.df.iloc[index.row(),index.column()] = value
            self.layoutChanged.emit([QPersistentModelIndex(index)])
            return True

    def headerData(self, section: int, orientation: Qt.Orientation, role: Qt.ItemDataRole):
        if not role == Qt.ItemDataRole.DisplayRole or orientation == Qt.Orientation.Vertical:
            return
        
        headers = self.df.columns
        
        return headers[section]

    def flags(self, index: QModelIndex) -> Qt.ItemFlag:
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEditable
    
    def sort(self, Ncol, order):
        return

    def _sort(self, Ncol, order):
        try:
            # self.layoutAboutToBeChanged.emit()
            self.df.sort_values(self.df.columns[Ncol], ascending=order, inplace=True)
            self.layoutChanged.emit()
        except Exception as e:
            print(e)
    
    
    def sorted(self, ix: int, next_sort_state: Literal['Asc', 'Desc', 'None'] = None):
        col_name = self.df.columns[ix]

        # Determine next sorting state by current state
        if next_sort_state is None:
            # Clicked an unsorted column
            if ix != self.sorted_column_ix:
                next_sort_state = 'Asc'
            # Clicked a sorted column
            elif ix == self.sorted_column_ix and self.sort_state == 'Asc':
                next_sort_state = 'Desc'
            # Clicked a reverse sorted column - reset to sorted by index
            elif ix == self.sorted_column_ix:
                next_sort_state = 'None'

        if next_sort_state == 0:
            self.df = self.df.sort_values(col_name, ascending=True, kind='mergesort')
            self.sorted_column_name = self.df.columns[ix]
            self.sort_state = 'Asc'

        elif next_sort_state == 1:
            self.df = self.df.sort_values(col_name, ascending=False, kind='mergesort')
            self.sorted_column_name = self.df.columns[ix]
            self.sort_state = 'Desc'

        elif next_sort_state == 'None':
            self.df = self.df.sort_index(ascending=True, kind='mergesort')
            self.sorted_column_name = None
            self.sort_state = 'None'

        self.layoutChanged.emit()


    def getHeaders(self, min, max=None):
        if max is None:
            return self.headerData(min, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole)
        
        _headers = []
        for i in range(min,max):
            _headers.append(self.headerData(i, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole))

        return _headers
    
class PandasView(QTableView):
    """
    Extends the QtableView class to include basic spreadsheet functionality like copy/paste and sorting cols

    """
    def __init__(self, *args, **kwargs):
        super(PandasView, self).__init__(*args, **kwargs)
        self.installEventFilter(self)

    def eventFilter(self, source, event):
        #print(event.type())
        if event.type() == QEvent.KeyPress and event.matches(QKeySequence.Copy):
            self.copy_selection()
            return True
        elif event.type() == QEvent.KeyPress and event.matches(QKeySequence.Paste):
            self.paste_selection()
            return True
        elif event.type() == QEvent.KeyPress and event.matches(QKeySequence.Delete):
            self.delete_selection()
            return True
        # elif event.type() == QMouseEvent and event.matches(QMouseEvent.mouseDoubleClickEvent):
        #     self.sort_selection()
        #     return True
        return super(PandasView, self).eventFilter(source, event)

    def delete_selection(self):
        selection = self.selectedIndexes()

        if not selection:
            return
        
        cols = []
        for index in selection:
            if not index.column() in cols:
                cols.append(index.column())
        
        model = self.model()

    def sort_selection(self):
        selection = self.selectedIndexes()

        if not selection:
            return
        
        cols = []
        for index in selection:
            if not index.column() in cols:
                cols.append(index.column())
        
        model = self.model()
        model._sort(cols[0], 1)

    def copy_selection(self):
        selection = self.selectedIndexes()

        if not selection:
            return
        
        all_rows = []
        all_columns = []
        headers = []
        for index in selection:
            if not index.row() in all_rows:
                all_rows.append(index.row())
            if not index.column() in all_columns:
                all_columns.append(index.column())
                headers.append(self.model().getHeaders(index.column()))

        visible_rows = [row for row in all_rows if not self.isRowHidden(row)]
        visible_columns = [col for col in all_columns if not self.isColumnHidden(col)]

        table = [[""] * len(visible_columns) for _ in range(len(visible_rows))]

        cols = []
        for index in selection:
            if index.row() in visible_rows and index.column() in visible_columns:
                if not index.column() in cols:
                    cols.append(index.column())
                selection_row = visible_rows.index(index.row())
                selection_column = visible_columns.index(index.column())
                data = index.data()
                if data == 'nan' or data == 'NaN':
                    data = ''
                table[selection_row][selection_column] = data
            
        col_check = False
        for col in cols:
            col_check = self.selectionModel().isColumnSelected(col, parent = QModelIndex())
        if col_check:
            table = [headers] + table


        stream = io.StringIO()
        csv.writer(stream, delimiter="\t").writerows(table)
        QApplication.clipboard().setText(stream.getvalue())

    def paste_selection(self):
        selection = self.selectedIndexes()
        if selection:
            model = self.model()

            buffer = QApplication.clipboard().text()
            all_rows = []
            all_columns = []
            for index in selection:
                if not index.row() in all_rows:
                    all_rows.append(index.row())
                if not index.column() in all_columns:
                    all_columns.append(index.column())
            visible_rows = [row for row in all_rows if not self.isRowHidden(row)]
            visible_columns = [
                col for col in all_columns if not self.isColumnHidden(col)
            ]

            reader = csv.reader(io.StringIO(buffer), delimiter="\t")
            arr = [[cell for cell in row] for row in reader]
            if len(arr) > 0: #there is something to paste
                nrows = len(arr)
                ncols = len(arr[0])
                justPasteItAll = True
                if len(visible_rows) == 1 and len(visible_columns) == 1 and not justPasteItAll:
                    # Only one cell highlighted.
                    for i in range(nrows):
                        insert_rows = [visible_rows[0]]
                        row = insert_rows[0] + 1
                        while len(insert_rows) < nrows:
                            row += 1
                            if not self.isRowHidden(row):
                                insert_rows.append(row)                              
                    for j in range(ncols):
                        insert_columns = [visible_columns[0]]
                        col = insert_columns[0] + 1
                        while len(insert_columns) < ncols:
                            col += 1
                            if not self.isColumnHidden(col):
                                insert_columns.append(col)
                    for i, insert_row in enumerate(insert_rows):
                        for j, insert_column in enumerate(insert_columns):
                            cell = arr[i][j]
                            model.setData(model.index(insert_row, insert_column), cell, Qt.EditRole)
                elif not justPasteItAll:
                    for index in selection:
                        selection_row = visible_rows.index(index.row())
                        selection_column = visible_columns.index(index.column())
                        try: 
                            model.setData(
                                model.index(index.row(), index.column()),
                                arr[selection_row][selection_column],
                                Qt.EditRole
                            )
                        except IndexError:
                            continue
                else:
                    topleftRow = visible_rows[0]
                    topleftCol = visible_columns[0]
                    for i in range(nrows):
                        for j in range(ncols):
                            print("Trying to set ", arr[i][j]," on row ", topleftRow+i," col ", topleftCol+j)
                            try: 
                                model.setData(
                                    model.index(topleftRow+i, topleftCol+j),
                                    arr[i][j],
                                    Qt.EditRole
                                )
                            except IndexError:
                                print("oops")
                                continue
                        

        return

    

'''old model for viewing only'''
# class pandasView(QAbstractTableModel):

#     def __init__(self, data):
#         QAbstractTableModel.__init__(self)
#         self.df = data

#     def rowCount(self, parent=None):
#         return self.df.shape[0]

#     def columnCount(self, parent=None):
#         return self.df.shape[1]

#     def data(self, index, role=Qt.DisplayRole):
#         if index.isValid():
#             if role == Qt.DisplayRole:
#                 return str(self.df.iloc[index.row(), index.column()])
#         return None

#     def headerData(self, col, orientation, role):
#         if orientation == Qt.Horizontal and role == Qt.DisplayRole:
#             return self.df.columns[col]
#         return None