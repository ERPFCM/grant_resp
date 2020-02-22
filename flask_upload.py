# -*- coding: utf-8 -*-
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import xlrd, cx_Oracle, config as cfg, pandas as pd, numpy as np, re, os, time
app = Flask(__name__)

#업로드 HTML 렌더링
@app.route('/upload')
def render_file(org=None, period=None, cost_type=None):
    org = request.args.get('org')
    period = request.args.get('period')
    cost_type = request.args.get('cost_type')
    # dict = {'organization':org, 'period':period, 'cost_type':cost_type}
    # print(dict)
    print(org + ' ' + period + ' ' + cost_type)
    # return render_template('upload.html', upload=dict)
    return render_template('upload.html', org=org, period=period, cost_type=cost_type)

#템플릿 다운로드
@app.route('/download')
def download_template():
    file_name = "business_plan_upload_template.xlsx"
    return send_file(file_name, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', attachment_filename='사업계획_업로드.xlsx', as_attachment=True)

#파일 업로드 처리
@app.route('/fileupload', methods = ['GET','POST'])
def upload_file():
    if request.method == 'POST':
        f = request.files['file']
        org = request.form['organization']
        period = request.form['period']
        cost_type = request.form['cost_type']
        #저장할 경로 + 파일명
        file_name = "D:/python/BusinessExpense/" + time.strftime('%Y%m%d%H%M%S', time.localtime()) + f.filename
        # file_name = "/usr/tmp/" + time.strftime('%Y%m%d%H%M%S', time.localtime()) + f.filename
        f.save(file_name)
        #f.save("D:/python/BusinessExpense/" + f.filename)
        print(file_name)
        print("ORG > " + org)
        # base_dir = 'D:/python/BusinessExpense'
        # excel_file = f.filename
        # excel_dir = os.path.join(base_dir, excel_file)
        # print(excel_dir)
        # 파일 내용을 html에 뿌려준다.(CSV)
        #df_to_html = pd.read_csv("D:/python/BusinessExpense/" + secure_filename(f.filename)).to_html()
        #return df_to_html
        os.putenv('NLS_LANG','.UTF8')
        print(time.strftime('%Y/%m/%d %H:%M:%S', time.localtime()))
        #conn = cx_Oracle.connect('APPS','devapps','TOBE_DEV')
        conn = cx_Oracle.connect(cfg.username, cfg.password, cfg.dsn, encoding = cfg.encoding)
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
        def item_check(t):
            sql_item_check = """
            select segment1
            from mtl_system_items_b msi
            , gmf_organization_definitions god
            where 1=1
            and msi.organization_id = god.organization_id
            and god.organization_code = :1
            """
            cursor.execute(sql_item_check,t)
            item_list = cursor.fetchall()
            return item_list
        def account_check(t):
            sql_account_check = """
            select xga.account_code
            from xcstf_gl_accounts_v xga
            , gmf_organization_definitions god
            where 1=1
            and xga.chart_of_accounts_id = god.chart_of_accounts_id
            and god.organization_code = :1
            """
            cursor.execute(sql_account_check,t)
            account_list = cursor.fetchall()
            return account_list
        def raise_error(sheet_name, line, value):
            cursor.close()
            conn.rollback()
            conn.close()
            error_msg = "에러가 발생하였습니다. \n" + sheet_name + " 의 " + str(line) + " 번 라인 데이터를 확인하세요.\n에러 발생 데이터 : " + value
            return error_msg
        # cursor.execute의 바인드 인자도 튜플이기 때문에 (org,) 와 같이 사용함.
        litem = item_check((org,))
        # DB조회 결과 튜플을 1차원 리스트로 변경.
        item_list = []
        for i in range(len(litem)):
            item_list.append(litem[i][0])
        #리스트 삭제
        del litem
        laccount = account_check((org,))
        account_list = []
        for i in range(len(laccount)):
            account_list.append(laccount[i][0])
        del laccount
        delete()
        sheet_name = 'Plan_Order'
        df = pd.read_excel(file_name, sheet_name=sheet_name, dtype={'ORG':str, 'ITEM':str, 'PERIOD':str, 'COST_TYPE':str, 'ORDER_QTY':str})
        for table in range(len(df)):
            row = df.iloc[table]
            if (row[0] != org):
                return raise_error(sheet_name, table + 2, row[0])
            elif (row[2] != period):
                return raise_error(sheet_name, table + 2, row[2])
            elif (row[3] != cost_type):
                return raise_error(sheet_name, table + 2, row[3])
            elif (row[1] not in item_list):
                return raise_error(sheet_name, table + 2, row[1])
            insert_order((row[0],row[1],row[2],row[3],row[4]))
        sheet_name = 'Plan_Expense'
        df = pd.read_excel(file_name, sheet_name=sheet_name, dtype={'ORG':str, 'COST_TYPE':str, 'PERIOD_CODE':str, 'PART_OF_FACTORY':str, 'ACCOUNT':str, 'EXPENSE_AMOUNT':str})
        for table in range(len(df)):
            row = df.iloc[table]
            if (row[0] != org):
                return raise_error(sheet_name, table + 2, row[0])
            elif (row[2] != period):
                return raise_error(sheet_name, table + 2, row[2])
            elif (row[1] != cost_type):
                return raise_error(sheet_name, table + 2, row[1])
            elif (row[4] not in account_list):
                return raise_error(sheet_name, table + 2, row[4])
            insert_expense((row[0], row[1], row[2], row[3], row[4], row[5]))
        sheet_name = 'Purchase_Unit_Cost'
        df = pd.read_excel(file_name, sheet_name=sheet_name, dtype={'ORG':str, 'ITEM':str, 'PERIOD':str, 'COST_TYPE':str, 'COMPONENT_CLASS':str, 'UNIT_COST':str})
        for table in range(len(df)):
            row = df.iloc[table]
            if (row[0] != org):
                return raise_error(sheet_name, table + 2, row[0])
            elif (row[2] != period):
                return raise_error(sheet_name, table + 2, row[2])
            elif (row[3] != cost_type):
                return raise_error(sheet_name, table + 2, row[3])
            elif (row[1] not in item_list):
                return raise_error(sheet_name, table + 2, row[1])
            insert_purchase((row[0], row[1], row[2], row[3], row[4], row[5]))
        sheet_name = 'Depreciation_Expense'
        df = pd.read_excel(file_name, sheet_name=sheet_name, dtype={'ORG':str, 'COST_TYPE':str, 'PERIOD':str, 'PART_OF_FACTORY':str, 'ACCOUNT':str, 'ALLOC_METHOD':str, 'EXPENSE_AMOUNT':str})
        for table in range(len(df)):
            row = df.iloc[table]
            if (row[0] != org):
                return raise_error(sheet_name, table + 2, row[0])
            elif (row[2] != period):
                return raise_error(sheet_name, table + 2, row[2])
            elif (row[1] != cost_type):
                return raise_error(sheet_name, table + 2, row[1])
            elif(row[4] not in account_list):
                return raise_error(sheet_name, table + 2, row[4])
            insert_depreciation((row[0], row[1], row[2], row[3], row[4], row[5], row[6]))
        cursor.close()
        conn.commit()
        conn.close()
        print(time.strftime('%Y/%m/%d %H:%M:%S', time.localtime()))

        return f.filename + ' 파일이 정상적으로 업로드 되었습니다.'

if __name__ == '__main__':
    #서버 실행
    app.run(debug = True)
    #app.run(host='0.0.0.0', port=5000)
