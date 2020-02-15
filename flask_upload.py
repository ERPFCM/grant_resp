from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
import xlrd, cx_Oracle, pandas as pd, numpy as np, re, os, time
app = Flask(__name__)

#업로드 HTML 렌더링
@app.route('/upload')
def render_file():
    return render_template('upload.html')

#파일 업로드 처리
@app.route('/fileupload', methods = ['GET','POST'])
def upload_file():
    if request.method == 'POST':
        f = request.files['file']
        #저장할 경로 + 파일명
        # file_name = "D:/python/BusinessExpense/" + secure_filename(f.filename)
        file_name = "D:/python/BusinessExpense/" + time.strftime('%Y%m%d%H%M%S', time.localtime()) + f.filename
        f.save(file_name)
        #f.save("D:/python/BusinessExpense/" + f.filename)
        print(file_name)
        # base_dir = 'D:/python/BusinessExpense'
        # excel_file = f.filename
        # excel_dir = os.path.join(base_dir, excel_file)
        # print(excel_dir)
        # 파일 내용을 html에 뿌려준다.(CSV)
        #df_to_html = pd.read_csv("D:/python/BusinessExpense/" + secure_filename(f.filename)).to_html()
        #return df_to_html
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
        df = pd.read_excel(file_name, sheet_name='판매계획', dtype={'ORG':str, 'Item':str, 'period':str, 'costmethod':str, 'Sales':str})
        # insert_stmt = 'insert into xcstf_upload_temp(attribute01, attribute02, attribute03, attribute04, attribute05, creation_date, program_id) '
        for table in range(len(df)):
            row = df.iloc[table]
            insert_order((row[0],row[1],row[2],row[3],row[4]))
        df = pd.read_excel(file_name, sheet_name='비용계획', dtype={'ORG_CODE':str, 'COST_TYPE':str, 'PERIOD_CODE':str, 'PART_OF_FACTORY':str, 'ACCOUNT':str, 'EXPENSE_AMOUNT':str})
        for table in range(len(df)):
            row = df.iloc[table]
            insert_expense((row[0], row[1], row[2], row[3], row[4], row[5]))
        df = pd.read_excel(file_name, sheet_name='구매단가', dtype={'조직':str, '품목':str, '기간':str, '원가유형':str, '원가요소':str, '금액':str})
        for table in range(len(df)):
            row = df.iloc[table]
            insert_purchase((row[0], row[1], row[2], row[3], row[4], row[5]))
        df = pd.read_excel(file_name, sheet_name='감가예산', dtype={'조직':str, '원가유형':str, '기간':str, '공장동구분':str, '계정':str, '배부기준':str, '예산':str})
        for table in range(len(df)):
            row = df.iloc[table]
            insert_depreciation((row[0], row[1], row[2], row[3], row[4], row[5], row[6]))
        cursor.close()
        conn.commit()
        conn.close()
        print(time.strftime('%Y/%m/%d %H:%M:%S', time.localtime()))

        return f.filename + ' 파일이 정상적으로 업로드 되었습니다.'

if __name__ == '__main__':
    #서버 실행
    app.run(debug = True)
