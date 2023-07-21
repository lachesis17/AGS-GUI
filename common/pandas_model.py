import pandas as pd
from PyQt5.QtCore import *
from PyQt5.QtWidgets import QApplication, QTableView
from PyQt5.QtGui import QKeySequence
import csv
import io

class PandasModel(QAbstractTableModel):
    def __init__(self, dataframe: pd.DataFrame):
        super().__init__()
        '''saving some commands to be used on subclass'''
        #self.installEventFilter(self)
        #self.table.setSortingEnabled(True)
        #self.table.horizontalHeader().sectionPressed.connect(self.table.selectColumn)
        
        self._dataframe = dataframe
        
    def rowCount(self, parent: QPersistentModelIndex) -> int:
        if self._dataframe is None:
            return 0
        else:
            return self._dataframe.shape[0]

    def columnCount(self, parent=QModelIndex) -> int:
        if self._dataframe is None:
            return 0
        else:
            return self._dataframe.shape[1]

    def data(self, index, role: int):
        if index.isValid():
            if role == Qt.DisplayRole or role == Qt.EditRole:
                return str(self._dataframe.iloc[index.row(), index.column()])
        return None

    def setData(self, index, value, role):
        try:
            value = float(value)
        except ValueError:
            pass
        
        if role == Qt.EditRole:
            self._dataframe.iloc[index.row(),index.column()] = value
            self.layoutChanged.emit([QPersistentModelIndex(index)])
            return True

    def headerData(self, section: int, orientation: Qt.Orientation, role: Qt.ItemDataRole):
        if not role == Qt.ItemDataRole.DisplayRole or orientation == Qt.Orientation.Vertical:
            return
        
        headers = self._dataframe.columns
        
        return headers[section]

    def flags(self, index: QModelIndex) -> Qt.ItemFlag:
        return Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEditable

    def sort(self, Ncol, order):
        """Sort table by column number."""
        try:
            self.layoutAboutToBeChanged.emit()
            self._dataframe.sort_values(self._dataframe.columns[Ncol], ascending=order, inplace=True)
            self.layoutChanged.emit()
        except Exception as e:
            print(e)

    def getHeaders(self, min, max=None):
        if max is None:
            return self.headerData(min, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole)
        
        _headers = []
        for i in range(min,max):
            _headers.append(self.headerData(i, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole))

        return _headers
    
class TableView(QTableView):
    """
    Extends the QtableView class to include basic spreadsheet functionality like copy/paste and sorting cols

    """
    def __init__(self, *args, **kwargs):
        super(TableView, self).__init__(*args, **kwargs)
        self.installEventFilter(self)

    def eventFilter(self, source, event):
        if event.type() == QEvent.KeyPress and event.matches(QKeySequence.Copy):
            self.copy_selection()
            return True
        elif event.type() == QEvent.KeyPress and event.matches(QKeySequence.Paste):
            self.paste_selection()
            return True
        return super(TableView, self).eventFilter(source, event)

    def copy_selection(self):
        selection = self.selectedIndexes()

        if not selection:
            return

        # This is just getting rid of duplicate indexes
        # Also I'm collecting the headers here
        all_rows = []
        all_columns = []
        headers = []
        for index in selection:
            if not index.row() in all_rows:
                all_rows.append(index.row())
            if not index.column() in all_columns:
                all_columns.append(index.column())
                headers.append(self.model().getHeaders(index.column()))

        # Keep rows and cols if they're not hidden
        visible_rows = [row for row in all_rows if not self.isRowHidden(row)]
        visible_columns = [col for col in all_columns if not self.isColumnHidden(col)]

        # Make a list of lists with empty strings for each cell
        table = [[""] * len(visible_columns) for _ in range(len(visible_rows))]

        col_check = False
        for index in selection:
            col_check = self.selectionModel().isColumnSelected(index.column(), parent = QModelIndex())
            if index.row() in visible_rows and index.column() in visible_columns:
                selection_row = visible_rows.index(index.row())
                selection_column = visible_columns.index(index.column())
                data = index.data()
                if data == 'nan' or data == 'NaN':
                    data = ''
                table[selection_row][selection_column] = data

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
#         self._data = data

#     def rowCount(self, parent=None):
#         return self._data.shape[0]

#     def columnCount(self, parent=None):
#         return self._data.shape[1]

#     def data(self, index, role=Qt.DisplayRole):
#         if index.isValid():
#             if role == Qt.DisplayRole:
#                 return str(self._data.iloc[index.row(), index.column()])
#         return None

#     def headerData(self, col, orientation, role):
#         if orientation == Qt.Horizontal and role == Qt.DisplayRole:
#             return self._data.columns[col]
#         return None