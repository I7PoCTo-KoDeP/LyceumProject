from PyQt5 import uic
from PyQt5.QtWidgets import QMainWindow, QTableWidgetItem, QAction, QWidget, QMenu, QMessageBox

ACHIEVEMENT = 1
LEARNING = 0
MULTIPLIER = 100
PER_GOOD_STUDENT = 30
PER_PERFECT_STUDENT = 50
ACHIEVEMENT_HEADER = ['Вид', 'Тип', 'Уровень', 'Место', 'Описание', 'Участники', 'Очки']
LEARNING_HEADER = ['Критерии', 'Кол-во', 'Ср.балл, кол-во', 'Очки']
LEARNING_CRITERIA = ['Ср. Балл по успеваемости', 'Отличники', 'Ударники']


class WorkspaceWindow(QMainWindow):
    def __init__(self, date, aspect, database):
        super().__init__()
        self.date = date
        self.aspect = aspect
        self.database = database
        self.is_changed = False
        if aspect == LEARNING:
            uic.loadUi('UIs/learning_workspace.ui', self)
            self.setFixedSize(620, 390)
        else:
            uic.loadUi('UIs/workspace.ui', self)
            self.setFixedSize(880, 700)
            self.initUI()
        self.load_table(self.aspect)

    def initUI(self):
        self.setWindowTitle(f'{self.date} - Рабочая область')
        menubar = self.menuBar()
        self.statusbar = self.statusBar()
        # Actions.
        cancel_action = QAction('&Отменить', self)
        cancel_action.triggered.connect(self.cancel_action)
        find_action = QAction('&Найти', self)
        find_action.triggered.connect(self.open_find_window)
        delete_action = QAction('Удалить выбранную строку', self)
        delete_action.triggered.connect(self.delete_row)
        # Menus.
        redo_menu = menubar.addMenu('&Правка')

        redo_menu.addAction(find_action)
        redo_menu.addAction(delete_action)
        redo_menu.addAction(cancel_action)
        self.add_button.clicked.connect(self.put_in_table)
        self.load_combo_boxes()

    def load_table(self, aspect):                                           # Вывод таблицы.
        num_of_students = 0                                                 # Также можно используется, как обновление
        total = 0                                                           # данных в таблице.
        points = 0
        res = self.database.get_table(self.date, self.aspect)
        self.data_output.setRowCount(0)
        if aspect == ACHIEVEMENT:
            if not res:
                return
            self.data_output.setColumnCount(len(res[0]) - 1)
            self.data_output.setHorizontalHeaderLabels(ACHIEVEMENT_HEADER)
            for i, row in enumerate(res):
                self.data_output.setRowCount(self.data_output.rowCount() + 1)
                for j, elem in enumerate(row):
                    if j == 7:
                        self.data_output.item(i, 4).setData(20, elem)
                    elif j == 5:
                        num_of_students = len(elem.split(', '))
                        self.data_output.setItem(i, j, QTableWidgetItem(str(elem)))
                    elif j == 6:
                        self.data_output.setItem(i, j, QTableWidgetItem(str(elem * num_of_students)))
                        total += elem * num_of_students
                    else:
                        self.data_output.setItem(i, j, QTableWidgetItem(str(elem)))
        else:
            res = res[0]
            self.data_output.setColumnCount(len(LEARNING_HEADER))
            self.data_output.setHorizontalHeaderLabels(LEARNING_HEADER)
            for i in range(len(LEARNING_CRITERIA)):
                self.data_output.setRowCount(self.data_output.rowCount() + 1)
                for j in range(len(LEARNING_HEADER)):
                    if j == 0:
                        self.data_output.setItem(i, j, QTableWidgetItem(str(LEARNING_CRITERIA[i])))
                    elif j == 1 and i > 0:
                        self.data_output.setItem(i, j, QTableWidgetItem(str(res[i])))
                        if res[i] != '':
                            num_of_students = len(res[i].split(', '))
                        else:
                            num_of_students = 0
                        if i == 1:
                            points = num_of_students * PER_PERFECT_STUDENT
                        elif i == 2:
                            points = num_of_students * PER_GOOD_STUDENT
                        total += points
                    elif j == 2 and i == 0:
                        self.data_output.setItem(i, j, QTableWidgetItem(str(res[i])))
                        points = float(res[i]) * MULTIPLIER
                        total += points
                    elif j == 2 and i > 0:
                        self.data_output.setItem(i, j, QTableWidgetItem(str(num_of_students)))
                    elif j == 3 and i > 0:
                        self.data_output.setItem(i, j, QTableWidgetItem(str(points)))
                    elif j == 3 and i == 0:
                        self.data_output.setItem(i, j, QTableWidgetItem(str(int(points))))
        self.data_output.setRowCount(self.data_output.rowCount() + 1)
        self.data_output.setItem(self.data_output.rowCount() - 1, 0, QTableWidgetItem('Итого'))
        self.data_output.setItem(self.data_output.rowCount() - 1, 1, QTableWidgetItem(str(int(total))))
        self.data_output.resizeColumnsToContents()
        self.data_output.cellChanged.connect(self.change_data)

    def load_combo_boxes(self):                                         # Загрузка интерфейсов comboBoxes.
        type_data = set(self.database.get_data('''SELECT Name FROM Type'''))
        aspect_data = self.database.get_data('''SELECT Name FROM Aspects''')
        level_data = self.database.get_data('''SELECT Name FROM Levels''')
        place_data = self.database.get_data('''SELECT Name FROM Places''')
        for i in aspect_data:
            self.aspect_box.addItem(*i)
        for i in type_data:
            self.type_box.addItem(*i)
        for i in place_data:
            self.place_box.addItem(*i)
        for i in level_data:
            self.level_box.addItem(*i)
        self.aspect_box.textActivated.connect(self.change_types)        # При изменении значения в одном ComboBox
        self.type_box.textActivated.connect(self.change_place)          # меняются значения и в других.
        self.type_box.textActivated.connect(self.change_level)

    def change_types(self, text):
        self.type_box.clear()
        aspect_id = self.database.get_data('''SELECT Id FROM Aspects WHERE Name = ?''', (text,))
        type_data = set(self.database.get_data('''SELECT Name FROM Type WHERE AspectId = ?''', *aspect_id))
        for i in type_data:
            self.type_box.addItem(*i)

    def change_place(self, text):
        self.place_box.clear()
        checklist = []
        place_ids = self.database.get_data('''SELECT PlaceId FROM Type WHERE Name = ?''', (text,))
        for i in place_ids:
            place = self.database.get_data('''SELECT Name FROM Places WHERE Id = ?''', i)[0]
            if place[0] not in checklist:
                checklist.append(*place)
                self.place_box.addItem(*place)

    def change_level(self, text):
        self.level_box.clear()
        checklist = []
        level_ids = self.database.get_data('''SELECT LevelId FROM Type WHERE Name = ?''', (text,))
        for i in level_ids:
            level = self.database.get_data('''SELECT Name FROM Levels WHERE Id = ?''', i)
            if level:
                level = level[0][0]
                if level not in checklist:
                    checklist.append(level)
                    self.level_box.addItem(level)
            else:
                return

    def put_in_table(self):                                             # Добавление ряда в базу данных.
        self.is_changed = True                                          # Работает только с достижениями
        if self.aspect == ACHIEVEMENT:
            request = [self.aspect_box.currentText(), self.type_box.currentText(), self.place_box.currentText(),
                       self.level_box.currentText(), self.request_edit.toPlainText()]
            self.database.add_to_database(self.aspect, self.date, *request)
            self.load_table(self.aspect)

    def change_data(self, row):                                       # Изменение данных.
        self.is_changed = True                                        # При изменении ячейки
        if self.aspect == ACHIEVEMENT:
            aspect = self.data_output.item(row, 0).text()
            _type = self.data_output.item(row, 1).text()
            level = self.data_output.item(row, 2).text()
            place = self.data_output.item(row, 3).text()
            _id = self.data_output.item(row, 4).data(20)
            name = self.data_output.item(row, 4).text()
            participants = self.data_output.item(row, 5).text()
            self.database.edit_data(self.aspect, self.date, _id, aspect, _type, level, place, name, participants)
        else:
            avg_points = self.data_output.item(0, 2).text()
            perf_stud = self.data_output.item(1, 1).text()
            good_stud = self.data_output.item(2, 1).text()
            self.database.edit_data(self.aspect, self.date, avg_points, perf_stud, good_stud)

    def delete_row(self):
        self.is_changed = True
        _id = self.data_output.item(self.data_output.currentRow(), 4).data(20)
        self.database.delete_data(_id)
        self.load_table(self.aspect)

    def cancel_action(self):
        self.database.undo_changes()
        self.load_table(self.aspect)

    def open_find_window(self):
        self.find_window = FindWindow()
        self.find_window.show()

    def closeEvent(self, event):
        if self.is_changed:
            valid = QMessageBox.question(self, 'Сохранение', 'Сохранить изменения?', QMessageBox.Yes, QMessageBox.No)
            if valid == QMessageBox.Yes:
                self.database.confirm_changes()


class FindWindow(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi('UIs/Info.ui', self)
        self.setFixedSize(460, 380)
