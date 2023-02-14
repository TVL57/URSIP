import openpyxl as op
import sqlite3 as sq
import datetime as dtt

# Принять Дата2 = текущая дата (но если Дата2.день = '01', то принять на день раньше текущей)
#         Дата1 = Дата2 - 1 день.

if dtt.date.today().strftime("%d") == '01':
    dt4 = dtt.date.today() - dtt.timedelta(days=1)
else:
    dt4 = dtt.date.today()

dt = dtt.timedelta(days=1)
dt3 = dt4 - dt
dt1=dt3.strftime("%Y%m%d")
dt2=dt4.strftime("%Y%m%d")

flxl = 'TB.xlsx'
wb = op.load_workbook(flxl, read_only=True)
sheet = wb.active
max_r = sheet.max_row

db = sq.connect('exdb.db')
crs = db.cursor()
print(crs.fetchone())

crs.execute("""CREATE TABLE IF NOT EXISTS tbl (
 id integer,
 company text,
 dat1 text,
 dat2 text,
 fact_qliq_dt1 integer,
 fact_qliq_dt2 integer,
 fact_qoil_dt1 integer,
 fact_qoil_dt2 integer,
 fore_qliq_dt1 integer,
 fore_qliq_dt2 integer,
 fore_qoil_dt1 integer,
 fore_qoil_dt2 integer
)""")

crs.execute("""DELETE FROM tbl""")

for row in range(4, max_r+1):
    val_row = [
        sheet[row][0].value,
        sheet[row][1].value,
        dt1,
        dt2,
        sheet[row][2].value,
        sheet[row][3].value,
        sheet[row][4].value,
        sheet[row][5].value,
        sheet[row][6].value, 
        sheet[row][7].value,
        sheet[row][8].value,
        sheet[row][9].value]
    crs.execute("INSERT INTO tbl VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", val_row)
    print(val_row)

crs.execute("""SELECT 
    dat1,
    sum(fact_qliq_dt1),
    sum(fact_qliq_dt2),
    sum(fact_qoil_dt1),
    sum(fact_qoil_dt2),
    sum(fore_qliq_dt1),
    sum(fore_qliq_dt2),
    sum(fore_qoil_dt1),
    sum(fore_qoil_dt2) FROM tbl GROUP BY dat1""")
print(crs.fetchone())
       
db.commit()
crs.close() 
db.close() 
