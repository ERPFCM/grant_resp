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
        def insert_allocation_item(t):
            sql_insert = "insert into xcstf_upload_temp(attribute01, attribute02, attribute03, attribute04, attribute05, attribute06, attribute07, attribute08, attribute09, attribute10, creation_date, program_id) values (:1, :2, :3, :4, :5, :6, :7, :8, :9, :10, sysdate, 'XCSTFF9050')"
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
        sl = pd.ExcelFile(file_name)
        sheet_name = 'Plan_Order'
        if sheet_name in sl.sheet_names:
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
            cursor.callproc("XCSTF_BUSINESS_PLAN_PKG.UPLOAD_PLAN_ORDER", [org, period, cost_type])
        sheet_name = 'Plan_Expense'
        if sheet_name in sl.sheet_names:
            df = pd.read_excel(file_name, sheet_name=sheet_name, dtype={'ORG':str, 'COST_TYPE':str, 'PERIOD':str, 'PART_OF_FACTORY':str, 'ACCOUNT':str, 'EXPENSE_AMOUNT':str})
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
            cursor.callproc("XCSTF_BUSINESS_PLAN_PKG.UPLOAD_PLAN_EXPENSE", [org, period, cost_type])
        sheet_name = 'Purchase_Unit_Cost'
        if sheet_name in sl.sheet_names:
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
            cursor.callproc("XCSTF_BUSINESS_PLAN_PKG.UPLOAD_MATERIAL_COST", [org, period, cost_type])
        sheet_name = 'Depreciation_Expense'
        if sheet_name in sl.sheet_names:
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
            cursor.callproc("XCSTF_BUSINESS_PLAN_PKG.UPLOAD_PLAN_DEP_EXPENSE", [org, period, cost_type])
        sheet_name = 'Allocation_Items'
        if sheet_name in sl.sheet_names:
            df = pd.read_excel(file_name, sheet_name=sheet_name,
                               dtype={'ORG': str, 'COST_TYPE': str, 'PERIOD': str, 'ITEM': str, 'WEIGHT': str, 'VARIABLE_CONSUMPTION': str, 'FIXED_LABOR': str, 'MANUFACTURE_EXPENSE': str, 'EMPLOYEE_BENEFIT': str, 'DEPRECIATION': str})
            chk_na = {'VARIABLE_CONSUMPTION':'N','FIXED_LABOR':'N','MANUFACTURE_EXPENSE':'N','EMPLOYEE_BENEFIT':'N','DEPRECIATION':'N'}
            df = df.fillna(chk_na)
            for table in range(len(df)):
                row = df.iloc[table]
                if (row[0] != org):
                    return raise_error(sheet_name, table + 2, row[0])
                elif (row[2] != period):
                    return raise_error(sheet_name, table + 2, row[2])
                elif (row[1] != cost_type):
                    return raise_error(sheet_name, table + 2, row[1])
                elif (row[3] not in item_list):
                    return raise_error(sheet_name, table + 2, row[3])
                elif df.isnull().loc[table,'WEIGHT']:
                    return raise_error(sheet_name, table + 2, 'Value is null')
                insert_allocation_item((row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9]))
            # cursor.callfunc("XCSTF_BUSINESS_PLAN_PKG.UPLOAD_ALLOC_ITEM", str, [t[0], t[2], t[1]])
            cursor.callproc("XCSTF_BUSINESS_PLAN_PKG.UPLOAD_ALLOC_ITEM", [org, period, cost_type])
        cursor.close()
        conn.commit()
        conn.close()
        print(time.strftime('%Y/%m/%d %H:%M:%S', time.localtime()))

        return f.filename + ' 파일이 정상적으로 업로드 되었습니다.'
# @app.route('/verify')
# def verify():
#     conn = cx_Oracle.connect(cfg.username, cfg.password, cfg.dsn, encoding=cfg.encoding)
#     cursor = conn.cursor()
#     org_lov = "select organization_code from gmf_organization_definitions where organization_code like 'P%' order by organization_code"
#     cost_type_lov = "select cost_mthd_code from  cm_mthd_mst"
#     period_lov = "select distinct period_code from cm_cldr_dtl where start_date >= to_date('202001','YYYYMM') order by period_code"
#     cursor.execute(org_lov)
#     org_list = cursor.fetchall()
#     cursor.execute(cost_type_lov)
#     cost_type_list = cursor.fetchall()
#     cursor.execute(period_lov)
#     period_list = cursor.fetchall()
#     print(org_list)
#     return render_template('verify.html', org_list=org_list, cost_type_list=cost_type_list, period_list=period_list)
@app.route('/verify', methods = ['GET','POST'])
# def extract(org=None, period=None, cost_type=None):
def verify():
    conn = cx_Oracle.connect(cfg.username, cfg.password, cfg.dsn, encoding=cfg.encoding)
    cursor = conn.cursor()
    org_lov = "select organization_code from gmf_organization_definitions where organization_code like 'P%' order by organization_code"
    cost_type_lov = "select cost_mthd_code from  cm_mthd_mst"
    period_lov = "select distinct period_code from cm_cldr_dtl where start_date >= to_date('202001','YYYYMM') order by period_code"
    calendar_lov = "select calendar_code from cm_cldr_hdr"
    cursor.execute(org_lov)
    org_list = cursor.fetchall()
    cursor.execute(cost_type_lov)
    cost_type_list = cursor.fetchall()
    cursor.execute(period_lov)
    period_list = cursor.fetchall()
    cursor.execute(calendar_lov)
    calendar_list = cursor.fetchall()
    # org = request.args.get('org')
    # period = request.args.get('period')
    # cost_type = request.args.get('cost_type')
    org = request.form.get('org_select')
    period = request.form.get('period_select')
    cost_type = request.form.get('cost_type_select')
    calendar = request.form.get('calendar_select')
    # dict = {'organization':org, 'period':period, 'cost_type':cost_type}
    # print(dict)
    # print(org + ' ' + period + ' ' + cost_type)
    conn = cx_Oracle.connect(cfg.username, cfg.password, cfg.dsn, encoding=cfg.encoding)
    cursor = conn.cursor()
    # print(cursor)
    verifysql="""
        select decode(org, null, '합계', org) org
         , item
         , description
         , gl_cls
         , inventory_item_id
         , quantity1
         , ind_qty
         , std_cost
         , c1
         , c2
         , c3
         , c4
         , c5
         , c6
         , c7
         , c8
         , c9
         , sum(c3s) c3s
         , sum(c4s) c4s
         , sum(c5s) c5s
         , sum(c6s) c6s
         , sum(c7s) c7s
         , sum(c8s) c8s
         , sum(c9s) c9s
        from   (select org
                     , item
                     , description
                     , gl_cls
                     , inventory_item_id
                     , quantity1
                     , ind_qty
                     , max(c1) c1
                     , max(c2) c2
                     , max(c3) c3
                     , max(c4) c4
                     , max(c5) c5
                     , max(c6) c6
                     , max(c7) c7
                     , max(c8) c8
                     , max(c9) c9
                     , nvl(max(c1),0)+nvl(max(c2),0)+nvl(max(c3),0)+nvl(max(c4),0)+nvl(max(c5),0)+nvl(max(c6),0)+nvl(max(c7),0)+nvl(max(c8),0)+nvl(max(c9),0) std_cost
                     , max(c3s) c3s
                     , max(c4s) c4s
                     , max(c5s) c5s
                     , max(c6s) c6s
                     , max(c7s) c7s
                     , max(c8s) c8s
                     , max(c9s) c9s
                from   (select god.organization_code org
                             , msi.segment1 item
                             , msi.description description
                             , mc.description gl_cls
                             , msi.inventory_item_id
                             , xpq.quantity1
                             , pmq.quantity1 ind_qty
                             , decode(ccd.cost_cmpntcls_id, 1, sum(ccd.cmpnt_cost)  , null) as c1
                             , decode(ccd.cost_cmpntcls_id, 2, sum(ccd.cmpnt_cost)  , null) as c2
                             , decode(ccd.cost_cmpntcls_id, 3, sum(ccd.cmpnt_cost)  , null) as c3
                             , decode(ccd.cost_cmpntcls_id, 4, sum(ccd.cmpnt_cost)  , null) as c4
                             , decode(cbd.cost_cmpntcls_id, 5, cbd.burden_usage     , null) as c5
                             , decode(cbd.cost_cmpntcls_id, 6, cbd.burden_usage     , null) as c6
                             , decode(cbd.cost_cmpntcls_id, 7, cbd.burden_usage     , null) as c7
                             , decode(cbd.cost_cmpntcls_id, 8, cbd.burden_usage     , null) as c8
                             , decode(cbd.cost_cmpntcls_id, 9, cbd.burden_usage     , null) as c9
                             , decode(ccd.cost_cmpntcls_id, 3, sum(ccd.cmpnt_cost)  , null)*xpq.quantity1 as c3s
                             , decode(ccd.cost_cmpntcls_id, 4, sum(ccd.cmpnt_cost)  , null)*xpq.quantity1 as c4s
                             , decode(cbd.cost_cmpntcls_id, 5, cbd.burden_usage     , null)*pmq.quantity1 as c5s
                             , decode(cbd.cost_cmpntcls_id, 6, cbd.burden_usage     , null)*pmq.quantity1 as c6s
                             , decode(cbd.cost_cmpntcls_id, 7, cbd.burden_usage     , null)*pmq.quantity1 as c7s
                             , decode(cbd.cost_cmpntcls_id, 8, cbd.burden_usage     , null)*pmq.quantity1 as c8s
                             , decode(cbd.cost_cmpntcls_id, 9, cbd.burden_usage     , null)*pmq.quantity1 as c9s
                        from   cm_cmpt_dtl ccd
                             , mtl_system_items_b msi
                             , gmf_organization_definitions god
                             , mtl_categories mc
                             , mtl_item_categories mic
                             , xcstf_plan_order_quantities xpq
                             , xcstf_business_plan_mfg_qty_v pmq
                             , cm_brdn_dtl cbd
                             , gmf_period_statuses gps
                             , cm_mthd_mst cmm
                        where  1=1
                        and    god.organization_id = mic.organization_id
                        and    mc.category_id = mic.category_id
                        and    mic.category_set_id = xcstf_category_set_id
                        and    mic.inventory_item_id = msi.inventory_item_id
                        and    god.organization_id = msi.organization_id
                        and    msi.organization_id = ccd.organization_id
                        and    msi.inventory_item_id = ccd.inventory_item_id
                        and    gps.cost_type_id = cmm.cost_type_id
                        and    gps.period_id = ccd.period_id(+) -- 2001 BPST
                        and    gps.cost_type_id = ccd.cost_type_id(+)
                        and    ccd.period_id = pmq.period_id(+)
                        and    ccd.cost_type_id = pmq.cost_type_id(+)
                        and    mc.segment1 in ('05','06')
                        and    god.organization_id = xpq.organization_id(+)
                        and    msi.inventory_item_id = xpq.inventory_item_id(+)
                        and    ccd.period_id = xpq.period_id(+)
                        and    god.organization_id = pmq.organization_id(+)
                        and    msi.inventory_item_id = pmq.inventory_item_id(+)
                        and    pmq.organization_id = cbd.organization_id(+)
                        and    pmq.inventory_item_id = cbd.inventory_item_id(+)
                        and    pmq.period_id = cbd.period_id(+)
                        and    pmq.cost_type_id = cbd.cost_type_id(+)
                        and    god.organization_code = :1
                        and    gps.period_code = :2
                        and    cmm.cost_mthd_code = :3
                        and    gps.calendar_code = :4
                        group by god.organization_code
                             , msi.segment1
                             , msi.description
                             , msi.inventory_item_id
                             , mc.description
                             , xpq.quantity1
                             , pmq.quantity1
                             , ccd.cost_cmpntcls_id
                             , cbd.cost_cmpntcls_id
                             , cbd.burden_usage )
                group by org
                     , item
                     , description
                     , gl_cls
                     , inventory_item_id
                     , quantity1
                     , ind_qty )
        group by rollup(( org
                             , item
                             , description
                             , gl_cls
                             , inventory_item_id
                             , quantity1
                             , ind_qty
                             , std_cost
                             , c1
                             , c2
                             , c3
                             , c4
                             , c5
                             , c6
                             , c7
                             , c8
                             , c9 ))
        union all
        select '예산'
             , null
             , null
             , null
             , null
             , null
             , null
             , null
             , null
             , null
             , null
             , null
             , null
             , null
             , null
             , null
             , null
             , max(decode(cost_cmpntcls_id, 3, expense_amount, 0))
             , max(decode(cost_cmpntcls_id, 4, expense_amount, 0))
             , max(decode(cost_cmpntcls_id, 5, expense_amount, 0))
             , max(decode(cost_cmpntcls_id, 6, expense_amount, 0))
             , max(decode(cost_cmpntcls_id, 7, expense_amount, 0))
             , max(decode(cost_cmpntcls_id, 8, expense_amount, 0))
             , max(decode(cost_cmpntcls_id, 9, expense_amount, 0))
        from   (select pev.organization_id
                     , pev.cost_cmpntcls_id
                     , sum(expense_amount) expense_amount
                from   xcstf_plan_expenses_v pev
                     , gmf_organization_definitions god
                where  1=1
                and    pev.organization_id = god.organization_id
                and    god.organization_code = :1
                and    pev.period_code = :2
                and    pev.cost_mthd_code = :3
                and    pev.calendar_code = :4
                group by pev.organization_id
                     , pev.cost_cmpntcls_id) 
    """
    t = (org, period, cost_type, calendar)
    cursor.execute(verifysql, t)
    verifydata = cursor.fetchall()
    print(t)
    return render_template('verify.html', org_list=org_list, cost_type_list=cost_type_list, period_list=period_list, calendar_list=calendar_list, data_list = verifydata)

if __name__ == '__main__':
    #서버 실행
    app.run(debug = True)
    #app.run(host='0.0.0.0', port=5000)

