#installing the libraries psycopg2, Workbook,pandas
import psycopg2
from openpyxl.workbook import Workbook
import pandas as pd

class Total_compensation:
    def compensation(self):
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
            query_data_command = """select emp.ename, emp.empno, dept.dname, (case when enddate is not null then ((enddate-startdate+1)/30)*(jobhist.sal) else ((current_date-startdate+1)/30)*(jobhist.sal) end)as Total_Compensation,
(case when enddate is not null then ((enddate-startdate+1)/30) else ((current_date-startdate+1)/30) end)as Months_Spent from jobhist, dept, emp 
where jobhist.deptno=dept.deptno and jobhist.empno=emp.empno"""
            cur.execute(query_data_command)
            columns = [desc[0] for desc in cur.description]
            data = cur.fetchall()
            dataframe = pd.DataFrame(list(data), columns=columns)
            # storing values inside excel
            writer = pd.ExcelWriter('ques_2.xlsx')
            #converting data frame to excel
            dataframe.to_excel(writer, sheet_name='bar')
            writer.save()

        except Exception as e:
            print("Something went wrong", e)
        finally:

            if con is not None:
                cur.close()
                con.close()


if __name__=='__main__':
    con = None
    cur = None
    comp = Total_compensation()
    comp.compensation()