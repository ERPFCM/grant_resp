import xlrd, cx_Oracle, pandas as pd, numpy as np, re, os, time
os.putenv('NLS_LANG','.UTF8')
print(time.strftime('%Y/%m/%d %H:%M:%S', time.localtime()))
conn = cx_Oracle.connect('APPS','devapps','TOBE_DEV')
cursor = conn.cursor()
def insert_order(t):
    sql_insert = "insert into xcstf_upload_temp(attribute01, attribute02, attribute03, attribute04, attribute05, creation_date, program_id) values (:1, :2, :3, :4, :5, sysdate, 'XCSTFF9020')"
    cursor.execute(sql_insert,t)
def insert_expense(t):
    sql_insert = "insert into xcstf_upload_temp(attribute01, attribute02, attribute03, attribute04, attribute05, attribute06, creation_date, program_id) values (:1, :2, :3, :4, :5, :6, sysdate, 'XCSTFF9030')"
    cursor.execute(sql_insert,t)
def insert_purchase(t):
    sql_insert = "insert into xcstf_upload_temp(attribute01, attribute02, attribute03, attribute04, attribute05, attribute06, creation_date, program_id) values (:1, :2, :3, :4, :5, :6, sysdate, 'XCSTFF9010')"
    cursor.execute(sql_insert,t)
def insert_depreciation(t):
    sql_insert = "insert into xcstf_upload_temp(attribute01, attribute02, attribute03, attribute04, attribute05, attribute06, attribute07, creation_date, program_id) values (:1, :2, :3, :4, :5, :6, :7, sysdate, 'XCSTFF9040')"
    cursor.execute(sql_insert,t)
def delete():
    sql_delete = "delete from xcstf_upload_temp"
    cursor.execute(sql_delete)
delete()
df = pd.read_excel('file_name', sheet_name='판매계획', dtype={'ORG':str, 'Item':str, 'period':str, 'costmethod':str, 'Sales':str})
# insert_stmt = 'insert into xcstf_upload_temp(attribute01, attribute02, attribute03, attribute04, attribute05, creation_date, program_id) '
for table in range(len(df)):
    row = df.iloc[table]
    insert_order((row[0],row[1],row[2],row[3],row[4]))
df = pd.read_excel('file_name', sheet_name='비용계획', dtype={'ORG_CODE':str, 'COST_TYPE':str, 'PERIOD_CODE':str, 'PART_OF_FACTORY':str, 'ACCOUNT':str, 'EXPENSE_AMOUNT':str})
for table in range(len(df)):
    row = df.iloc[table]
    insert_expense((row[0], row[1], row[2], row[3], row[4], row[5]))
df = pd.read_excel('file_name', sheet_name='구매단가', dtype={'조직':str, '품목':str, '기간':str, '원가유형':str, '원가요소':str, '금액':str})
for table in range(len(df)):
    row = df.iloc[table]
    insert_purchase((row[0], row[1], row[2], row[3], row[4], row[5]))
df = pd.read_excel('file_name', sheet_name='감가예산', dtype={'조직':str, '원가유형':str, '기간':str, '공장동구분':str, '계정':str, '배부기준':str, '예산':str})
for table in range(len(df)):
    row = df.iloc[table]
    insert_depreciation((row[0], row[1], row[2], row[3], row[4], row[5], row[6]))
cursor.close()
conn.commit()
conn.close()
print(time.strftime('%Y/%m/%d %H:%M:%S', time.localtime()))