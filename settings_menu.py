from PyQt5 import uic
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QTableWidgetItem, QWidget, QFileDialog, QMessageBox
from constants import SETTINGS_HEADER


class SettingsMenu(QWidget):
    sig = pyqtSignal()

    def __init__(self, database):
        super().__init__()
        self.isChanged = False
        self.database = database
        uic.loadUi('appdata/UIs/settings_menu.ui', self)
        self.setWindowTitle('Настройки')
        self.setFixedSize(520, 440)
        self.initUI()

    def initUI(self):
        data = self.database.send_request('''SELECT * FROM Other_data''')[0]
        self.class_edit.setText(str(data[1]))
        self.num_of_class_edit.setText(data[2])
        self.edu_years_edit.setText(data[0])
        with open('appdata/settings_file.txt', mode='r', encoding='utf-8') as f:
            path = f.read()
            self.file_path_edit.setText(path)
        self.load_points_table()
        self.close_button.clicked.connect(self.close)
        self.save_button.clicked.connect(self.save)
        self.path_button.clicked.connect(self.file_path)

        self.class_edit.textChanged.connect(self.line_edit_changed)
        self.num_of_class_edit.textChanged.connect(self.line_edit_changed)
        self.edu_years_edit.textChanged.connect(self.line_edit_changed)

    def load_points_table(self):
        res = self.database.send_request('''SELECT
                                             Aspects.Name as AspectName,
                                             Type.Name as Type,
                                             Levels.Name as LevelName,
                                             Places.Name as Place,
                                             Type.Points
                                         FROM
                                             Type
                                         LEFT JOIN Aspects ON Type.AspectId = Aspects.Id
                                         LEFT JOIN Levels ON Type.LevelId = Levels.Id
                                         LEFT JOIN Places ON Type.PlaceId = Places.Id
                                         ORDER BY AspectName;''')
        self.tableWidget.setColumnCount(5)
        self.tableWidget.setRowCount(0)
        self.tableWidget.setHorizontalHeaderLabels(SETTINGS_HEADER)
        for i, row in enumerate(res):
            self.tableWidget.setRowCount(self.tableWidget.rowCount() + 1)
            for j, elem in enumerate(row):
                self.tableWidget.setItem(i, j, QTableWidgetItem(str(elem)))
        self.tableWidget.cellChanged.connect(self.change_scoring)
        self.tableWidget.resizeColumnsToContents()

    def save(self):                                 # Сохранение изменений настроек приложения и данных в БД.
        self.isChanged = False
        with open('appdata/settings_file.txt', mode='w', encoding='utf-8') as f:
            f.seek(0)
            f.write(self.file_path_edit.text())
        self.database.send_request('''UPDATE Other_data SET class = ?, num_of_class = ?, date = ?''',
                                   (int(self.class_edit.text()), self.num_of_class_edit.text(),
                                    self.edu_years_edit.text()))
        self.database.confirm_changes()
        self.close()

    def file_path(self):                            # Выбор пути к БД.
        f_name = QFileDialog.getOpenFileName(self, 'Выбрать путь к файлу', '', 'Таблицы (*.sqlite);;Все файлы (*)')[0]
        if f_name != '':
            self.isChanged = True
            self.file_path_edit.setText(f_name)

    def change_scoring(self, row):
        self.isChanged = True
        aspect = self.tableWidget.item(row, 0).text()
        name = self.tableWidget.item(row, 1).text()
        level = self.tableWidget.item(row, 2).text()
        place = self.tableWidget.item(row, 3).text()
        value = self.tableWidget.item(row, 4).text()
        self.database.send_request('''UPDATE Type SET Points = ? WHERE AspectId=(SELECT Id FROM Aspects WHERE Name = ?) 
                                   AND Name = ? AND LevelId=(SELECT Id FROM Levels WHERE Name = ?) AND
                                   PlaceId=(SELECT Id FROM Places WHERE Name = ?)''',
                                   (value, aspect, name, level, place))

    def line_edit_changed(self):
        self.isChanged = True

    def closeEvent(self, event):
        self.sig.emit()
        if self.isChanged:
            valid = QMessageBox.question(self, 'Сохранение', 'Сохранить изменения?', QMessageBox.Yes, QMessageBox.No)
            if valid == QMessageBox.Yes:
                self.save()
