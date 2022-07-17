import openpyxl
import psycopg2


host = "127.0.0.1"
user = "postgres"
password = "Sonas_cnizekov"
db_name = ""


try:
    connection = psycopg2.connect(
        host = host,
        user = user,
        password = password,
        database = db_name,
    )
    book = openpyxl.open("названия точек.xlsm", read_only=True)
    sheet = book.active

    #once create db
    # with connection.cursor() as cursor:
    #     cursor.execute(
    #         """CREATE TABLE task2(
    #             id serial PRIMARY KEY,
    #             endpoint_id numeric,
    #             endpoint_name varchar)"""
    #     )
    #     connection.commit()
    #     print('db created')
    with connection.cursor() as cursor:
        k=0
        for i in sheet:
           k+=1
           print (i[0].value, ' ', i[1].value)
           timer = i[1].value
           if(i[0].value == None):
               break
           cursor.execute(
                   f"""UPDATE task2
                SET endpoint_id = {i[0].value},
                endpoint_name = '{timer}'
                WHERE id = {k}
""")
        connection.commit()
except Exception as _ex:
    print(_ex)
finally:
    if connection:
        connection.close()
        print("connection closed")

