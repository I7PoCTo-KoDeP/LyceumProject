import sys
import database_script
import file_generator
import workspace
from PyQt5 import uic
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTableWidgetItem, QAction, QWidget, QMenu, QDialog,
                             QListWidgetItem)

ACHIEVEMENT = 1
LEARNING = 0
MULTIPLIER = 100
PER_GOOD_STUDENT = 30
PER_PERFECT_STUDENT = 50
DATABASE_PATH = 'data/auto-sys-database.sqlite'
ACHIEVEMENT_HEADER = ['Вид', 'Тип', 'Уровень', 'Место', 'Описание', 'Участники', 'Очки']
LEARNING_HEADER = ['Критерии', 'Кол-во', 'Ср.балл, кол-во', 'Очки']
LEARNING_CRITERIA = ['Ср. Балл по успеваемости', 'Отличники', 'Ударники']
MAIN_CRITERIA = ['Учёба', 'Внеурочная деятельность', 'Спорт', 'Активная жизненная позиция']


class AchievementControl(QMainWindow):
    def __init__(self):
        super().__init__()
        self.roster = {}
        self.database = database_script.Database(DATABASE_PATH)
        self.docx = file_generator.CreateWordFile(self.database)
        # Load UI.
        uic.loadUi('UIs/cadet.ui', self)
        self.setFixedSize(800, 470)
        self.initUI()

    def initUI(self):
        menubar = self.menuBar()
        # Actions.
        makefile_action = QAction('&Сгенерировать файл', self)
        makefile_action.triggered.connect(self.create_docx_file)
        save_action = QAction('&Сохранить', self)
        about_action = QAction('&О программе', self)
        about_action.triggered.connect(self.open_info)
        # Menus.
        file_menu = menubar.addMenu('&Файл')
        workspace_menu = menubar.addMenu('&Открыть таблицу')
        info_menu = menubar.addMenu('&Справка')
        # Submenus.
        self.learning_menu = QMenu('&Учёба', self)
        self.achievement_menu = QMenu('&Достижения', self)
        # Fill submenus.
        self.login()
        # Another methods.
        info_menu.addAction(about_action)
        file_menu.addAction(save_action)
        file_menu.addAction(makefile_action)
        workspace_menu.addMenu(self.learning_menu)
        workspace_menu.addMenu(self.achievement_menu)
        # Initialization methods.
        self.load_table()
        self.show_best_students()

    def quarters(self):
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

    def half_year(self):
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

    def login(self):
        with open('data/settings_file.txt', mode='r', encoding='utf-8') as f:
            if int(f.read(2)) >= 10:
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

    def open_workspace(self):
        self.w_space = workspace.WorkspaceWindow(self.sender().text()[self.sender().text().find(' ') + 1:],
                                                 self.sender().data(), self.database)
        self.w_space.show()

    def open_info(self):
        self.info_menu = Info()
        self.info_menu.show()

    def calculate_learning(self, date):
        result = 0
        learning = self.database.get_data('''SELECT * FROM Grade''')
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
        achievements = self.database.get_data('''SELECT
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
        learning = self.database.get_data('''SELECT * FROM Grade''')
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
        achievements = self.database.get_data('''SELECT MainTable.Participants, Type.Points FROM MainTable
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

    def create_docx_file(self):
        self.docx.docx_file_generator(self.n, self.tableWidget)

    def closeEvent(self, event):
        self.database.close_connection()


class RegistrationWindow(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi('UIs/register.ui', self)
        self.save_button.clicked.connect(self.save)

    def save(self):
        with open('data/settings_file.txt', mode='w', encoding='utf-8') as f:
            f.write(self.class_edit.text() + ' ' + self.platoon_edit.text())
            self.accept()
            self.close()


class Info(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi('UIs/Info.ui', self)
        self.setFixedSize(460, 380)
        self.close_button.clicked.connect(self.close)


class SettingsMenu(QWidget):
    def __init__(self):
        super().__init__()


def registered():
    try:
        with open('data/settings_file.txt', mode='r', encoding='utf-8') as f:
            if f.read(1) == '':
                return False
    except FileNotFoundError:
        return False
    return True


if __name__ == '__main__':
    app = QApplication(sys.argv)
    if not registered():
        login = RegistrationWindow()
        if login.exec_() == QDialog.Accepted:
            ex = AchievementControl()
            ex.show()
            sys.exit(app.exec_())
    else:
        ex = AchievementControl()
        ex.show()
        sys.exit(app.exec_())
