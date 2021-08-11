#installing the libraries psycopg2, Workbook,pandas
import psycopg2
from openpyxl.workbook import Workbook
import pandas as pd

class employee_info:
    def emp(self):
        # to connect to the PostgreSQL database server in the Python program using the psycopg database adapter.
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
