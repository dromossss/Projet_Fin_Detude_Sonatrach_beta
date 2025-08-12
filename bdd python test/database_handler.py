import os
import sqlite3

class DatabaseHandler:
    def __init__(self, database_name):
        self.database_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), database_name)
        self.con = sqlite3.connect(self.database_path)
        self.con.row_factory = sqlite3.Row
        
    def create_person_table(self):
        cursor = self.con.cursor()
        query = """
            CREATE TABLE IF NOT EXISTS Person (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL,
                password TEXT NOT NULL
            )
        """
        cursor.execute(query)
        self.con.commit()
        cursor.close()
    
    def create_person(self, username, password):
        self.create_person_table()
        cursor = self.con.cursor()
        query = "INSERT INTO Person (username, password) VALUES (?, ?)"
        cursor.execute(query, (username, password))
        self.con.commit()
        cursor.close()


# Example usage
if __name__ == "__main__":
    database_handler = DatabaseHandler("mydatabase.db")
    database_handler.create_person("john", "password123")
