#installing the libraries psycopg2, Workbook,pandas
import psycopg2
from openpyxl.workbook import Workbook
import pandas as pd

''' class not commented to describe it like
class for extracting employee information like employee name,employee no. , manager'''

class employee_info:
    def emp(self):
        # to connect to the PostgreSQL database server in the Python program using the psycopg database adapter.
        
''' connection to database is made everytime in each file this database connection would have been enclosed inside a class in main file 
   like:
   
   
   class Config_database:
    #function to connect the database
    def connect(self):
        try:
            self.connection = psycopg2.connect("dbname=postgres user=postgres password=password")
            self.cursor=self.connection.cursor()
            logging.info("connected to database")
            return self.cursor
        except Exception as e:
            logging.error("error in connecting to database")
            raise Exception(e)

    #function to commit the operations
    def commit(self):
        self.connection.commit()
        logging.info("Operations commited")
        self.connection.close()
        
        and this class function should be called each time in every file as Config_database.connect()
        '''
        
        try:
            con = psycopg2.connect(
                host="localhost",
                database="Assignment1",
                user="postgres",
                password="912313")
            # Creating a cursor object using the cursor() method
            cur = con.cursor()
            # Reading table which we imported using connection through query
            query_data_command = """SELECT e1.empno, e1.ename, (case when mgr is not null then (select ename from emp as e2 where e1.mgr=e2.empno limit 1) else null end) as manager
            from emp as e1"""
            cur.execute(query_data_command)

            columns = [desc[0] for desc in cur.description]
            data = cur.fetchall()
            dataframe = pd.DataFrame(list(data), columns=columns)
            # storing values inside excel
            writer = pd.ExcelWriter('ques_1.xlsx')
            # converting data frame to excel
            dataframe.to_excel(writer, sheet_name='bar')
            writer.save()

        except Exception as e:
            print("Something went wrong", e)
        finally:

            if con is not None:
                cur.close()
                con.close()


if __name__=='__main__':

    employee = employee_info()
    employee.emp()
