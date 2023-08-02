import pandas as pd
from PyQt5.QtCore import *
#from PyQt5.QtCore import QAbstractTableModel, QPersistentModelIndex, QModelIndex, QEvent, QTimer
from PyQt5.QtWidgets import QApplication, QTableView, QDoubleSpinBox, QMenu, QInputDialog, QPushButton
from PyQt5.QtGui import QKeySequence, QMouseEvent, QIcon, QPixmap
import PyQt5.QtCore as QtCore
import csv
import io
import webbrowser
import sys
sys.stdout.reconfigure(encoding='utf-8')

class PandasModel(QAbstractTableModel):
    def __init__(self, dataframe: pd.DataFrame):
        super().__init__()
        '''saving some commands to be used on subclass'''
        #self.installEventFilter(self)
        #self.table.setSortingEnabled(True)
        #self.table.horizontalHeader().sectionPressed.connect(self.table.selectColumn)
        
        self.original = dataframe.copy()
        self.df = dataframe
        self.sort_state = 0
        
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
                return self.df.iloc[index.row(), index.column()]
        return None

    def setData(self, index, value, role):
        try:
            value = float(value)
        except ValueError:
            pass
        
        if role == Qt.EditRole:
            self.df.iloc[index.row(),index.column()] = value
            self.layoutChanged.emit()
            return True

    # def headerData(self, section: int, orientation: Qt.Orientation, role: Qt.ItemDataRole):
    #     if not role == Qt.ItemDataRole.DisplayRole or orientation == Qt.Orientation.Vertical:
    #         return
        
    #     headers = self.df.columns
        
    #     return headers[section]

    def headerData(self, section: int, orientation: Qt.Orientation, role: Qt.ItemDataRole):
        # if not role == Qt.ItemDataRole.DisplayRole or orientation == Qt.Orientation.Vertical:
        #     return
        if role == Qt.ItemDataRole.DecorationRole and orientation == Qt.Orientation.Horizontal:
            if self.sort_state == 1:
                return QPixmap("common/images/sort-ascending.svg").scaled(100, 100, transformMode=QtCore.Qt.SmoothTransformation, aspectRatioMode=QtCore.Qt.KeepAspectRatio)
            elif self.sort_state == 2:
                return QPixmap("common/images/sort-descending.svg").scaled(100, 100, transformMode=QtCore.Qt.SmoothTransformation, aspectRatioMode=QtCore.Qt.KeepAspectRatio)
            else:
                pass
        elif not role == Qt.ItemDataRole.DisplayRole or orientation == Qt.Orientation.Vertical:
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


    def getHeaders(self, min, max=None):
        if max is None:
            return self.headerData(min, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole)
        
        _headers = []
        for i in range(min,max):
            _headers.append(self.headerData(i, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole))

        return _headers
    
    
class PandasView(QTableView):
    def __init__(self, *args, **kwargs):
        super(PandasView, self).__init__(*args, **kwargs)
        self.installEventFilter(self)
        self.headers = self.horizontalHeader()
        self.headers.sectionDoubleClicked.connect(lambda x: self.sort(x))
        self.headers.setContextMenuPolicy(Qt.CustomContextMenu)
        self.headers.customContextMenuRequested.connect(self.header_menu)


    def header_menu(self, position):
        menu = QMenu()
        model = self.model()
        rename = menu.addAction(QIcon("common/images/edit.svg"),"Rename Header")
        insert_col = menu.addAction(QIcon("common/images/insert.svg"),"Insert Column")
        del_col = menu.addAction(QIcon("common/images/delete.svg"),"Delete Column")
        move_right = menu.addAction(QIcon("common/images/right.svg"),"Move Column Right")
        move_left = menu.addAction(QIcon("common/images/left.svg"),"Move Column Left")
        sort_asc = menu.addAction(QIcon("common/images/sort-ascending.svg"),"Sort Ascending")
        sort_des = menu.addAction(QIcon("common/images/sort-descending.svg"),"Sort Descending")
        refresh = menu.addAction(QIcon("common/images/refresh.svg"),"Reload Data")
        menu.addSeparator()
        menu.addSeparator()
        github = menu.addAction(QIcon("common/images/github.svg"),"GitHub")
        _action = menu.exec_(self.mapToGlobal(position))
        index = self.headers.logicalIndexAt(position)
        col_name = model.df.columns[index]
        if _action == rename:
            new_header = QInputDialog.getText(self," ","New header name:")
            if new_header[1]:
                model.df.rename(columns={f'{col_name}':f'{new_header[0]}'}, inplace=True)
                self.resizeColumnsToContents()
                model.layoutChanged.emit()
        if _action == insert_col:   
            new_col = QInputDialog.getText(self," ","New column name:")
            if new_col[1]:
                try:
                    model.df.insert(index+1, new_col[0], value="")
                    model.layoutChanged.emit()
                    self.resizeColumnsToContents()
                except Exception as e:
                    print(e)
        if _action == del_col:
            model.df.drop(col_name, axis=1, inplace=True)
            model.layoutChanged.emit()
            self.resizeColumnsToContents()
        if _action == move_right:
            if index + 1 >= len(model.df.columns):
                return
            model.df.insert(index+1, col_name, model.df.pop(col_name))
            self.resizeColumnsToContents()
            model.layoutChanged.emit()
        if _action == move_left:
            if index - 1 < 1:
                return
            model.df.insert(index-1, col_name, model.df.pop(col_name))
            self.resizeColumnsToContents()
            model.layoutChanged.emit()
        if _action == sort_asc:
            model.sort_state = 1
            model.df.sort_values(col_name, ascending=True, kind='mergesort', inplace=True)
            model.headerData(index, Qt.Orientation.Horizontal, role=Qt.ItemDataRole.DecorationRole)
            self.resizeColumnsToContents()
            model.layoutChanged.emit()
        if _action == sort_des:
            model.sort_state = 2
            model.df.sort_values(col_name, ascending=False, kind='mergesort', inplace=True)
            model.headerData(index, Qt.Orientation.Horizontal, role=Qt.ItemDataRole.DecorationRole)
            self.resizeColumnsToContents()
            model.layoutChanged.emit()
        if _action == refresh:
            model.df = model.original.copy()
            model.layoutChanged.emit()
            self.resizeColumnsToContents()
            model.sort_state = 0
            model.headerData(index, Qt.Orientation.Horizontal, role=Qt.ItemDataRole.DecorationRole)
        if _action == github:
            webbrowser.open('https://github.com/lachesis17')


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
        return super(PandasView, self).eventFilter(source, event)
    

    def sort(self, idx):
        model = self.model()
        col_name = model.df.columns[idx]

        #https://gist.github.com/StephenNneji/14bfc4e7a322ec89df7d30847fbf19b3
        #https://stackoverflow.com/questions/65179468/cannot-set-header-data-with-qtableview-custom-table-model
        #https://forum.qt.io/topic/54115/text-and-icon-in-qtableview-header
    
#        Determine next sorting state by current state
        if model.sort_state == 0:
            model.sort_state = 1
            model.df.sort_values(col_name, ascending=True, kind='mergesort', inplace=True)
            model.headerData(idx, Qt.Orientation.Horizontal, role=Qt.ItemDataRole.DecorationRole)
            self.resizeColumnsToContents()
            model.layoutChanged.emit()
            return
        if model.sort_state == 1:
            model.sort_state = 2
            model.df.sort_values(col_name, ascending=False, kind='mergesort', inplace=True)
            model.headerData(idx, Qt.Orientation.Horizontal, role=Qt.ItemDataRole.DecorationRole)
            self.resizeColumnsToContents()
            model.layoutChanged.emit()
            return
        if model.sort_state == 2:
            model.sort_state = 0
            model.df.sort_index(ascending=True, kind='mergesort', inplace=True)
            model.headerData(idx, Qt.Orientation.Horizontal, role=Qt.ItemDataRole.DecorationRole)
            self.resizeColumnsToContents()
            model.layoutChanged.emit()
            return


    def delete_selection(self):
        selection = self.selectedIndexes()

        if not selection:
            return
        
        cols = []
        rows = []
        for index in selection:
            if not index.column() in cols:
                cols.append(index.column())
            if not index.row() in rows:
                rows.append(index.row())

        model = self.model()
        head = model.getHeaders(min=cols, max=None)
        sel_cols = list(head)

        #checking to see if entire rows/columns are selected. for columns, use column name to get index position for int to use as arg for isColumnSelected(), rows already have int index
        col_check = False
        row_check = False
        col_idx = []
        for col in sel_cols:
            col_idx.append(list(model.df).index(col))
        for col in col_idx:
            col_check = self.selectionModel().isColumnSelected(col, parent = QModelIndex())
            if col_check:
                break
        for row in rows:
            row_check = self.selectionModel().isRowSelected(row, parent = QModelIndex())
            if row_check:
                break
        if col_check:
            model.df.drop(sel_cols, axis=1, inplace=True)
        if row_check:
            model.df.drop(rows, axis=0, inplace=True)
            model.df.reset_index(drop=True, inplace=True) #make sure to reset the index, as the row indexes stay the same in selection, will crash if trying to delete same index twice without resetting

        if not col_check and not row_check: #deleting cells if the entire row or entire column is not selected
            for index in selection:
                model.df.iloc[index.row(), index.column()] = ""

        # idx_top = model.createIndex(0,0)
        # idx_bot = model.createIndex(model.df.shape[0],0)
        # model.dataChanged.emit(idx_top, idx_bot)
        model.layoutChanged.emit()
        self.clearSelection()
        #self.viewport().repaint()

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

    
class Spinny(QDoubleSpinBox):
    #handles the event of clicking (focusIn) in the QDoubleSpinBox for override values to select all text to make it easier to paste values

    #__pyqtSignals__ = ("valueChanged(double)", "FocusIn()")

    def __init__(self, *args):
        QDoubleSpinBox.__init__(self, *args)

    # def event(self, event):
    #     if(event.type()==QEvent.Enter):
    #         self.emit(SIGNAL("Enter()"))
    #         #self.clear() 
    #         self.selectAll()               

    #     return QDoubleSpinBox.event(self, event)

    def focusInEvent(self, event) -> None:
        if(event.type()==QEvent.FocusIn):
            QTimer.singleShot(0, self.selectAll)


class GitButton(QPushButton):
    def __init__(self, *args, **kwargs):
        QPushButton.__init__(self, *args, **kwargs)
        self.setMouseTracking(True)

    def mouseMoveEvent(self, event):
        
        if (event.type()==QEvent.Leave):
            print('pp')
            icon = QPixmap("common/images/github.svg").scaled(25, 25, transformMode=QtCore.Qt.SmoothTransformation, aspectRatioMode=QtCore.Qt.KeepAspectRatio)
            self.setIcon(QIcon(icon))
            return QPushButton.mouseMoveEvent(self, event)
        else:
            icon_ = QPixmap("common/images/github_grey.svg").scaled(25, 25, transformMode=QtCore.Qt.SmoothTransformation, aspectRatioMode=QtCore.Qt.KeepAspectRatio)
            self.setIcon(QIcon(icon_))
            return QPushButton.mouseMoveEvent(self, event)

# bool myWidget::event(QEvent* e) 
# {
#     if(e->type() == QEvent::Leave) 
#     {
#         QPoint view_pos(x(), y());
#         QPoint view_pos_global = mapToGlobal(view_pos);
#         QPoint mouse_global = QCursor::pos();
#         if(mouse_global.x() < view_pos_global.x() || mouse_global.x() > view_pos_global.x() + width())
#         {
#             closeMenu();
#         }
#         else if(mouse_global.y() < view_pos_global.y() || mouse_global.y() > view_pos_global.y() + height())
#         {
#             closeMenu();
#         }
#     }
#     return QWidget::event(e);
# }


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