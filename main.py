import mysql.connector
from excel import writeExcel
import os
import time

hostname = 'localhost'
username = 'root'
password = ''
database = 'googlesheet'
table = 'pricetable'

def main():
    conn = mysql.connector.connect(
      host=hostname,
      user=username,
      password=password,
      database=database
    )
    cursor = conn.cursor()

    cursor.execute("select * from " + table)

    result = cursor.fetchall()

    values = [('id', 'Full Name', 'NI', 'Wages', 'Mileage', 'Other Expences')]

    for row in result:
        values.append((row[0], row[1], row[3], row[4]*row[5]+row[6]*row[7]+row[8]*row[9], row[10], 
            row[11]*row[12]+row[13]*row[14]+row[15]*row[16]+row[17]*row[18]+row[19]*row[20]+row[11]*row[22]))
    # print(values)

    writeExcel(values)

    os.system('start output.xlsx')
    
if __name__ == '__main__':
    main()