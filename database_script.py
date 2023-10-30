import sqlite3

ACHIEVEMENT = 1
LEARNING = 0


class Database:
    def __init__(self, path):
        self.connection = sqlite3.connect(path)
        self.cursor = self.connection.cursor()

    def get_table(self, date, aspect):
        if aspect == ACHIEVEMENT:
            return self.cursor.execute('''SELECT
                                                    Aspects.Name as AspectName,
                                                    Type.Name as Type,
                                                    Levels.Name as LevelName,
                                                    Places.Name as Place,
                                                    MainTable.Name,
                                                    MainTable.Participants,
                                                    Type.Points,
                                                    MainTable.Id
                                                FROM
                                                    MainTable
                                                LEFT JOIN Aspects ON MainTable.AspectId = Aspects.Id
                                                LEFT JOIN Type ON MainTable.TypeId = Type.Id
                                                LEFT JOIN Levels ON Type.LevelId = Levels.Id
                                                LEFT JOIN Places ON Type.PlaceId = Places.Id
                                                WHERE Date = ?
                                                ORDER BY AspectName;''', (date,)).fetchall()
        else:
            return self.cursor.execute('''SELECT AverageScore, PerfectStudents, GoodStudents FROM Grade
                                            WHERE Date = ?''', (date,)).fetchall()

    def get_data(self, request, *args):
        if args:
            return self.cursor.execute(request, *args).fetchall()
        return self.cursor.execute(request).fetchall()

    def add_to_database(self, aspect, date, aspect_name, type_name, place_name, level_name, other):
        if aspect == ACHIEVEMENT:
            other = other.split('; ')
            place_id = self.cursor.execute('''SELECT Id FROM Places WHERE Name = ?''', (place_name,)).fetchone()
            level_id = self.cursor.execute('''SELECT Id FROM Levels WHERE Name = ?''', (level_name,)).fetchone()
            aspect_id = self.cursor.execute('''SELECT Id FROM Aspects WHERE Name = ?''', (aspect_name,)).fetchone()
            type_id = self.cursor.execute('''SELECT Id FROM Type WHERE Name = ? AND PlaceId = ? AND AspectId = ? AND 
                                          LevelId = ?''', (type_name, *place_id, *aspect_id, *level_id)).fetchone()
            self.cursor.execute('''INSERT INTO MainTable VALUES(?, ?, ?, ?, null, ?)''',
                                (*aspect_id, other[0], other[1], *type_id, date)).fetchone()

    def edit_data(self, aspect, date, *new_row):
        if aspect == ACHIEVEMENT:
            ach_id, aspect_name, type_name, level, place_name, name, stud = new_row
            level_id = self.cursor.execute('''SELECT Id FROM Levels WHERE Name = ?''', (level,)).fetchone()
            place_id = self.cursor.execute('''SELECT Id FROM Places WHERE Name = ?''', (place_name,)).fetchone()
            aspect_id = self.cursor.execute('''SELECT Id FROM Aspects WHERE Name = ?''', (aspect_name,)).fetchone()
            type_id = self.cursor.execute('''SELECT Id FROM Type WHERE Name = ? AND PlaceId = ? AND AspectId = ? AND 
                                          LevelId = ?''', (type_name, *place_id, *aspect_id, *level_id)).fetchone()
            self.cursor.execute('''UPDATE MainTable
                                SET AspectId = ?, Name = ?, Participants = ?, TypeId = ? WHERE Id = ?''',
                                (*aspect_id, name, stud, *type_id, ach_id)).fetchall()
        else:
            avg, perf, good = new_row
            self.cursor.execute('''UPDATE Grade
                                SET AverageScore = ?, PerfectStudents = ?, GoodStudents = ?
                                WHERE Date = ?''', (avg, perf, good, date)).fetchall()

    def delete_data(self, row_id):
        self.cursor.execute('''DELETE FROM MainTable WHERE id = ?''', (row_id,)).fetchall()

    def undo_changes(self):
        pass

    def confirm_changes(self):
        self.connection.commit()

    def close_connection(self):
        self.connection.close()
