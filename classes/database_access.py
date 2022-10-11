import pymysql

class DB_Connect():
    """A class for connecting to the database"""

    def __init__(self, passed_db_username, passed_db_password, passed_database):
        """Initalizes database connection"""
        self.passed_db_username = passed_db_username
        self.passed_db_password = passed_db_password
        self.passed_database = passed_database
        self.conn = None 

    def __connect(self):
        """Creates connection to the database when needed"""
        self.conn = pymysql.connect(host='192.168.50.200',
                        user=self.passed_db_username,
                        password=self.passed_db_password,
                        db=self.passed_database,
                        charset='utf8mb4',
                        cursorclass=pymysql.cursors.DictCursor)

    def executeQuery(self, passed_query):
        """Executes queries through the database"""

        if not self.conn:
            self.__connect()

        with self.conn.cursor() as cursor:
            cursor.execute(passed_query)

    def executeSelectQuery(self, passed_query):
        """Executes a SELECT query and returns the results"""

        if not self.conn:
            self.__connect()

        with self.conn.cursor() as cursor:
            cursor.execute(passed_query)

        return cursor.fetchall()

    def close_connection(self):
        """Closes the database connection"""

        self.conn.close()