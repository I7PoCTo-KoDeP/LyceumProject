import sys
import database_module
import file_generator
import workspace
import settings_menu
from PyQt5 import uic
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTableWidgetItem, QAction, QWidget, QMenu, QDialog,
                             QListWidgetItem, QFileDialog)
from constants import *


class AchievementControl(QMainWindow):
    def __init__(self):
        super().__init__()
        self.database = database_module.database_connection()
        self.roster = {}
        self.docx = file_generator.CreateWordFile(self.database)
        # Загрузка интерфейсов.
        uic.loadUi('UIs/cadet.ui', self)
        self.setWindowTitle('Система учёта достижений кадет')
        self.setFixedSize(800, 470)
        self.initUI()

    def initUI(self):
        self.statusbar = self.statusBar()
        menubar = self.menuBar()
        menubar.clear()
        # Действия.
        open_action = QAction('&Открыть...', self)
        open_action.triggered.connect(self.open_file)
        makefile_action = QAction('&Сгенерировать файл', self)
        makefile_action.triggered.connect(self.create_docx_file)
        settings_action = QAction('&Настройки', self)
        settings_action.triggered.connect(self.open_settings)
        about_action = QAction('&О программе', self)
        about_action.triggered.connect(self.open_info)
        # Меню.
        file_menu = menubar.addMenu('&Файл')
        workspace_menu = menubar.addMenu('&Открыть таблицу')
        info_menu = menubar.addMenu('&Справка')
        # Субменю.
        self.learning_menu = QMenu('&Учёба', self)
        self.achievement_menu = QMenu('&Достижения', self)
        # Заполнение субменю.
        self.login()
        # Другие методы.
        info_menu.addAction(about_action)
        file_menu.addAction(open_action)
        file_menu.addAction(makefile_action)
        file_menu.addSeparator()
        file_menu.addAction(settings_action)
        workspace_menu.addMenu(self.learning_menu)
        workspace_menu.addMenu(self.achievement_menu)
        # Инициализация.
        self.load_table()
        self.show_best_students()

    def quarters(self):                                                 # Четвертная система обучения.
        self.n = 4
        for i in range(1, 5):
            learning_action = QAction(f'&Открыть {i} четверть', self)
            achievement_action = QAction(f'&Открыть {i} четверть', self)
            learning_action.setData(LEARNING)
            achievement_action.setData(ACHIEVEMENT)
            learning_action.triggered.connect(self.open_workspace)
            achievement_action.triggered.connect(self.open_workspace)
            self.learning_menu.addAction(learning_action)
            self.achievement_menu.addAction(achievement_action)

    def half_year(self):                                                # Полугодовая система обучения
        self.n = 2
        for i in range(1, 3):
            learning_action = QAction(f'&Открыть {i} полугодие', self)
            achievement_action = QAction(f'&Открыть {i} полугодие', self)
            learning_action.setData(LEARNING)
            achievement_action.setData(ACHIEVEMENT)
            learning_action.triggered.connect(self.open_workspace)
            achievement_action.triggered.connect(self.open_workspace)
            self.learning_menu.addAction(learning_action)
            self.achievement_menu.addAction(achievement_action)

    def login(self):                                                    # Выбор системы обучения и промежутка обучения.
        data = self.database.send_request('''SELECT Class, Date FROM Other_data''')
        self.title.setText(f'Результаты за {data[0][1]} учебный год:')
        if data[0][0] >= 10:
            self.half_year()
        else:
            self.quarters()

    def load_table(self):
        total = 0
        total_for_piece = [0] * self.n
        self.tableWidget.setColumnCount(self.n + 1)
        self.tableWidget.setRowCount(4)
        if self.n == 2:
            self.tableWidget.setHorizontalHeaderLabels(['', '1 полугодие', '2 полугодие'])
        else:
            self.tableWidget.setHorizontalHeaderLabels(['', '1 четверть', '2 четверть', '3 четверть', '4 четверть'])
        for i in range(len(MAIN_CRITERIA)):
            for j in range(self.n + 1):
                points = 0
                if j == 0:
                    self.tableWidget.setItem(i, j, QTableWidgetItem(MAIN_CRITERIA[i]))
                if i == 0 and j != 0:
                    points = self.calculate_learning(j)
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(points)))
                if i == 1 and j != 0:
                    points = self.calculate_achievements(j, MAIN_CRITERIA[i])
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(points)))
                if i == 2 and j != 0:
                    points = self.calculate_achievements(j, MAIN_CRITERIA[i])
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(points)))
                if i == 3 and j != 0:
                    points = self.calculate_achievements(j, MAIN_CRITERIA[i])
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(points)))
                total_for_piece[j - 1] += points
                total += points
        self.tableWidget.setRowCount(self.tableWidget.rowCount() + 1)
        self.tableWidget.setItem(self.tableWidget.rowCount() - 1, 0, QTableWidgetItem('Итого'))
        for i, val in enumerate(total_for_piece):
            self.tableWidget.setItem(self.tableWidget.rowCount() - 1, i + 1, QTableWidgetItem(str(int(val))))
        self.tableWidget.setRowCount(self.tableWidget.rowCount() + 1)
        self.tableWidget.setItem(self.tableWidget.rowCount() - 1, 0, QTableWidgetItem('За учебный год'))
        self.tableWidget.setItem(self.tableWidget.rowCount() - 1, 1, QTableWidgetItem(str(int(total))))
        self.tableWidget.resizeColumnsToContents()
        for i in range(self.tableWidget.rowCount()):
            if self.tableWidget.item(i, 0).text() == 'Спорт':
                self.paint_table(i, QColor(120, 240, 120))
            elif self.tableWidget.item(i, 0).text() == 'Учёба':
                self.paint_table(i, QColor(120, 120, 240))
            elif self.tableWidget.item(i, 0).text() == 'Внеурочная деятельность':
                self.paint_table(i, QColor(255, 184, 65))
            elif self.tableWidget.item(i, 0).text() == 'Активная жизненная позиция':
                self.paint_table(i, QColor(240, 120, 120))

    def paint_table(self, row, color):
        for i in range(self.tableWidget.columnCount()):
            self.tableWidget.item(row, i).setBackground(color)

    def calculate_learning(self, date):                                 # Подсчёт очков обучения.
        result = 0
        learning = self.database.send_request('''SELECT * FROM Grade''')
        for avg_score, perf_stud, good_stud, dt in learning:
            if int(dt[0]) == date:
                n, k = len(perf_stud.split(', ')), len(good_stud.split(', '))
                if perf_stud == '':
                    n = 0
                if good_stud == '':
                    k = 0
                result = int(avg_score * MULTIPLIER + n * PER_PERFECT_STUDENT + k * PER_GOOD_STUDENT)
        if result is None:
            return 0
        return result

    def calculate_achievements(self, date, aspect):
        result = 0
        achievements = self.database.send_request('''SELECT
                                              Aspects.Name as AspectName,
                                              MainTable.Participants,
                                              Type.Points,
                                              MainTable.Date
                                              FROM
                                              MainTable
                                              LEFT JOIN Aspects ON MainTable.AspectId = Aspects.Id
                                              LEFT JOIN Type ON MainTable.TypeId = Type.Id
                                              ORDER BY AspectName;''')
        for asp, participant, points, dt in achievements:
            if asp == aspect and int(dt[0]) == date:
                result += len(participant.split(', ')) * points
        if result is None:
            return 0
        return result

    def show_best_students(self):
        self.listWidget.clear()
        self.roster = {}
        learning = self.database.send_request('''SELECT * FROM Grade''')
        for i in learning:
            for j in i[1].split(', '):
                if j not in self.roster:
                    self.roster[j] = PER_PERFECT_STUDENT
                else:
                    self.roster[j] += PER_PERFECT_STUDENT
            for j in i[2].split(', '):
                if j not in self.roster:
                    self.roster[j] = PER_GOOD_STUDENT
                else:
                    self.roster[j] += PER_GOOD_STUDENT
        achievements = self.database.send_request('''SELECT MainTable.Participants, Type.Points FROM MainTable
                                                  LEFT JOIN Type ON MainTable.TypeId = Type.Id''')
        for i in achievements:
            for j in i[0].split(', '):
                if j in self.roster:
                    self.roster[j] += i[1]
                else:
                    self.roster[j] = i[1]
        sorted_roster = sorted(self.roster.items(), key=lambda x: x[1], reverse=True)
        for i, j in enumerate(sorted_roster):
            if j[0] != '':
                self.listWidget.addItem(QListWidgetItem(j[0].ljust(15, ' ') + str(j[1])))

    def sync(self):
        self.load_table()
        self.show_best_students()

    def create_docx_file(self):
        self.docx.docx_file_generator(self.n, self.tableWidget)

    def open_file(self):
        try:
            f_name = QFileDialog.getOpenFileName(self, 'Выбрать путь к файлу', '',
                                                 'Таблицы (*.sqlite);;Все файлы (*)')[0]
            if f_name == '':
                return
            self.database = database_module.Database(f_name)
            self.load_table()
            self.show_best_students()
            self.docx = file_generator.CreateWordFile(self.database)
        except BaseException:
            self.statusbar.showMessage('Неподдерживаемый формат файла, пожалуйста, выберите другой.')

    def open_workspace(self):
        try:
            self.w_space = workspace.WorkspaceWindow(self.sender().text()[self.sender().text().find(' ') + 1:],
                                                     self.sender().data(), self.database)
            self.w_space.show()
            self.w_space.sig.connect(self.sync)
        except BaseException as e:
            self.statusbar.showMessage('Непредвиденная ошибка: ' + str(e))

    def open_info(self):
        self.info_menu = Info()
        self.info_menu.show()

    def open_settings(self):
        self.settings = settings_menu.SettingsMenu(self.database)
        self.settings.show()
        self.settings.sig.connect(self.initUI)

    def closeEvent(self, event):
        self.database.close_connection()


class RegistrationWindow(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi('UIs/register.ui', self)
        self.save_button.clicked.connect(self.save)
        self.setFixedSize(270, 130)

    def save(self):
        database = database_module.database_connection()
        database.send_request('''INSERT INTO Other_data VALUES(?, ?, ?)''',
                              (self.edu_year_edit.text(), int(self.class_edit.text()), self.platoon_edit.text()))
        database.confirm_changes()
        self.accept()


class Info(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi('UIs/Info.ui', self)
        self.setWindowTitle('Информация о приложении')
        self.setFixedSize(460, 380)
        self.close_button.clicked.connect(self.close)


def registered():
    database = database_module.database_connection()
    data = database.send_request('''SELECT * FROM Other_data''')
    if not data or data[0][0] is None:
        return False
    return True


if __name__ == '__main__':
    app = QApplication(sys.argv)
    if not registered():
        login = RegistrationWindow()
        if login.exec() == QDialog.Accepted:
            ex = AchievementControl()
            ex.show()
            sys.exit(app.exec())
    else:
        ex = AchievementControl()
        ex.show()
        sys.exit(app.exec())
