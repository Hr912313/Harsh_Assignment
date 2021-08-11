

import psycopg2
from openpyxl.workbook import Workbook
import pandas as pd


class Employees:

    def emp(self):
        try:
            # read the connection parameters
            # connect to the PostgreSQL server
            conn = psycopg2.connect(
                host="localhost",
                database="Assignment1",
                user="postgres",
                password="912313")
            cursor = conn.cursor()
            #query for required result
            script = """
                    select dept.deptno as Dept_No, Total_Compensation.dname as Dept_Name, sum(total_compensation) as Compensation from Total_Compensation, dept
                    where Total_Compensation.dname=dept.dname
                    group by Total_Compensation.dname, dept.deptno
                    """
            # executing result
            cursor.execute(script)

            columns = [desc[0] for desc in cursor.description]
            data = cursor.fetchall()
            df = pd.DataFrame(list(data), columns=columns)
            # writing to xlsx file
            writer = pd.ExcelWriter('Ques_4.xlsx')
            df.to_excel(writer, sheet_name='bar')
            writer.save()

        except Exception as e:
            print("Error", e)

        finally:

            if conn is not None:
                cursor.close()
                conn.close()


if __name__ == '__main__':
    conn = None
    cur = None
    employee = Employees() # creating object of class
    employee.emp() # method calling
