# -*- coding: utf-8 -*-
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import xlrd, cx_Oracle, config as cfg, pandas as pd, numpy as np, re, os, time, openpyxl
app = Flask(__name__)
os_path_prefix = "D:/python/BusinessExpense/"
# os_path_prefix = "/usr/tmp/"
# export_file_name = ''
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

@app.route('/export')
def export_file():
    conn = cx_Oracle.connect(cfg.username, cfg.password, cfg.dsn, encoding=cfg.encoding)
    export_data = pd.read_sql(verifysql, conn, None, True, verify_param)
    export_file_name = "Export" + time.strftime('%Y%m%d%H%M%S', time.localtime()) + ".xlsx"
    # global export_file
    export_file = os.path.join(os_path_prefix, export_file_name)
    export_data.to_excel(export_file,  # directory and file name to write
                         sheet_name="Sheet1",
                         na_rep='',
                         float_format="%.9f",
                         header=True,
                         # columns = ["group", "value_1", "value_2"], # if header is False
                         index=False,
                         # index_label="id",
                         startrow=1,
                         startcol=1,
                         # engine = 'xlsxwriter',
                         freeze_panes=(2, 0)
                         )
    # print(export_data)
    file_name = export_file
    # print(file_name)
    conn.close()
    return send_file(file_name, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', attachment_filename=file_name, as_attachment=True)

#파일 업로드 처리
@app.route('/fileupload', methods = ['GET','POST'])
def upload_file():
    if request.method == 'POST':
        f = request.files['file']
        org = request.form['organization']
        period = request.form['period']
        cost_type = request.form['cost_type']
        #저장할 경로 + 파일명
        file_name = os_path_prefix + time.strftime('%Y%m%d%H%M%S', time.localtime()) + f.filename
        # file_name = os_path_prefix + time.strftime('%Y%m%d%H%M%S', time.localtime()) + f.filename
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
            # print(t)
            sql_insert = "insert into xcstf_upload_temp(attribute01, attribute02, attribute03, attribute04, attribute05, attribute06, creation_date, program_id) values (:1, :2, :3, :4, :5, :6, sysdate, 'XCSTFF9020')"
            # sql_insert2 = "insert into xcstf_upload_temp(attribute01, attribute02, attribute03, attribute04, attribute05, attribute06, creation_date, program_id) values (1, 2, 3, 4, 5, sysdate, 'XCSTFF9020')"
            # print(sql_insert2)
            cursor.execute(sql_insert,t)
            # cursor.execute(sql_insert2)
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
            conn.commit()
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
            select xga.cost_account
            from xcstf_cost_accounts xga
            , gmf_organization_definitions god
            where 1=1
            and xga.chart_of_accounts_id = god.chart_of_accounts_id
            and god.organization_code = :1
            """
            cursor.execute(sql_account_check,t)
            account_list = cursor.fetchall()
            return account_list
        def dep_account_check(t):
            sql_account_check = """
            select xga.cost_account
            from xcstf_cost_accounts xga
            , gmf_organization_definitions god
            where 1=1
            and xga.chart_of_accounts_id = god.chart_of_accounts_id
            and xga.cost_cmpntcls_id = 9
            and god.organization_code = :1
            """
            cursor.execute(sql_account_check,t)
            dep_account_list = cursor.fetchall()
            return dep_account_list
        def raise_resource_weights_error(sheet_name):
            cursor.close()
            conn.rollback()
            conn.close()
            error_msg = "에러가 발생하였습니다. \n " + sheet_name + " 는 1라인만 입력되어야 합니다. \n 데이터를 확인하여 주세요."
            return error_msg
        def raise_item_duplicated_error(sheet_name, item):
            cursor.close()
            conn.rollback()
            conn.close()
            error_msg = "에러가 발생하였습니다. \n " + sheet_name + " 의 " + item + " 이 중복데이터가 존재합니다."
            return error_msg
        def raise_error(sheet_name, line, value):
            cursor.close()
            conn.rollback()
            conn.close()
            error_msg = "에러가 발생하였습니다. \n " + sheet_name + " 의 " + str(line) + " 번 라인 데이터를 확인하세요.\n에러 발생 데이터 : " + value
            return error_msg
        def raise_dep_account_error(sheet_name, line, value):
            cursor.close()
            conn.rollback()
            conn.close()
            error_msg = "에러가 발생하였습니다. \n " + sheet_name + " 의 " + str(line) + " 번 라인 데이터를 확인하세요.\n에러 발생 데이터 : " + value
            return error_msg
        def raise_not_dep_account_error(sheet_name, line, value):
            cursor.close()
            conn.rollback()
            conn.close()
            error_msg = "에러가 발생하였습니다. 감가상각비는 Depreciation_Expense에 입력해주세요. \n " + sheet_name + " 의 " + str(line) + " 번 라인 데이터를 확인하세요.\n에러 발생 데이터 : " + value
            return error_msg
        def raise_execute_procedure_error(sheet_name, err_msg):
            cursor.close()
            conn.rollback()
            conn.close()
            error_msg = "Error occurred while uploading excel sheet (" + sheet_name + ") : " + err_msg
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
        daccount = dep_account_check((org,))
        dep_account_list = []
        for i in range(len(daccount)):
            dep_account_list.append(daccount[i][0])
        del daccount
        delete()
        sl = pd.ExcelFile(file_name)
        sheet_name = 'Plan_Order'
        if sheet_name in sl.sheet_names:
            print("Reading Plan_Order worksheet....")
            df = pd.read_excel(file_name, sheet_name=sheet_name, dtype={'ORG':str, 'ITEM':str, 'PERIOD':str, 'COST_TYPE':str, 'ORDER_QTY':str, 'PART_OF_FACTORY':str})
            # print(df)
            # Item duplication check > DataFrame을 List로 변환하여 check
            dup_check = df.duplicated('ITEM').to_list()
            item_list = df['ITEM'].to_list()
            for rec in range(len(item_list)):
                if dup_check[rec] == True:
                    return raise_item_duplicated_error(sheet_name, item_list[rec])
            # Data validation
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
                insert_order((row[0],row[1],row[2],row[3],row[4],row[5]))
            # conn.commit()
            print("Calling xcstf_business_plan_pkg.upload_plan_order..");
            outVal = cursor.var(str)
            cursor.callproc("XCSTF_BUSINESS_PLAN_PKG.UPLOAD_PLAN_ORDER", [org, period, cost_type, outVal])
            if outVal.getvalue() is not None:
                return raise_execute_procedure_error(sheet_name, outVal.getvalue())
        sheet_name = 'Plan_Expense'
        print("Reading Plan_Expense worksheet....")
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
                elif (row[4] in dep_account_list):
                    return raise_not_dep_account_error(sheet_name, table + 2, row[4])
                insert_expense((row[0], row[1], row[2], row[3], row[4], row[5]))
            outVal = cursor.var(str)
            cursor.callproc("XCSTF_BUSINESS_PLAN_PKG.UPLOAD_PLAN_EXPENSE", [org, period, cost_type, outVal])
            if outVal.getvalue() is not None:
                return raise_execute_procedure_error(sheet_name, outVal.getvalue())
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
            outVal = cursor.var(str)
            cursor.callproc("XCSTF_BUSINESS_PLAN_PKG.UPLOAD_MATERIAL_COST", [org, period, cost_type, outVal])
            if outVal.getvalue() is not None:
                return raise_execute_procedure_error(sheet_name, outVal.getvalue())
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
                elif(row[4] not in dep_account_list):
                    return raise_dep_account_error(sheet_name, table + 2, row[4])
                insert_depreciation((row[0], row[1], row[2], row[3], row[4], row[5], row[6]))
            outVal = cursor.var(str)
            cursor.callproc("XCSTF_BUSINESS_PLAN_PKG.UPLOAD_PLAN_DEP_EXPENSE", [org, period, cost_type, outVal])
            if outVal.getvalue() is not None:
                return raise_execute_procedure_error(sheet_name, outVal.getvalue())
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
            conn.commit()
            outVal = cursor.var(str)
            cursor.callproc("XCSTF_BUSINESS_PLAN_PKG.UPLOAD_ALLOC_ITEM", [org, period, cost_type, outVal])
            if outVal.getvalue() is not None:
                return raise_execute_procedure_error(sheet_name, outVal.getvalue())
        sheet_name = 'Copy_Resource_Weights'
        if sheet_name in sl.sheet_names:
            df = pd.read_excel(file_name, sheet_name=sheet_name,
                               dtype={'ORG': str, 'COPY_FROM_COST_TYPE': str, 'COPY_FROM_PERIOD': str, 'COPY_TO_COST_TYPE': str, 'COPY_TO_PERIOD': str})
            if (len(df) > 1):
                return raise_resource_weights_error(sheet_name)
            for table in range(len(df)):
                resource_weight_parameters = df.iloc[table]
                cursor.callproc("XCSTF_BUSINESS_PLAN_PKG.UPLOAD_RESOURCE_WEIGHT", resource_weight_parameters)
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
    # conn = cx_Oracle.connect(cfg.username, cfg.password, cfg.dsn, encoding=cfg.encoding)
    # cursor = conn.cursor()
    # print(cursor)
    global verifysql
    verifysql="""
        select decode(org, null, '합계', org) org
         , item
         , description
         , gl_cls
         , inventory_item_id
         , quantity1 "계획수량(가공비)"
         , ind_qty "계획수량(간접비)"
         , std_cost "표준원가"
         , c1 "원재료비"
         , c2 "포장재비"
         , c3 "직접노무"
         , c4 "기계가동"
         , c5 "변동소모"
         , c6 "고정노무"
         , c7 "생산경비"
         , c8 "복리후생"
         , c9 "감가상각"
         , sum(c3s) "직접노무금액"
         , sum(c4s) "기계가동금액"
         , sum(c5s) "변동소모금액"
         , sum(c6s) "고정노무금액"
         , sum(c7s) "생산경비금액"
         , sum(c8s) "복리후생금액"
         , sum(c9s) "감가상각금액"
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
                     , pev.cost_cmpntcls_id
                union all
                select pde.organization_id, ccm.cost_cmpntcls_id, sum(pde.expense_amount) expense_amount 
                from xcstf_plan_dep_expenses pde
                , gmf_organization_definitions god
                , gmf_period_statuses gps
                , cm_mthd_mst cmm
                , cm_cmpt_mst ccm
                where  1=1
                and    pde.organization_id = god.organization_id
                and    pde.period_id = gps.period_id
                and    cmm.cost_type_id = pde.cost_type_id
                and    pde.cost_type_id = gps.cost_type_id
                and    god.organization_code = :1
                and    gps.period_code = :2
                and    cmm.cost_mthd_code = :3
                and    gps.calendar_code = :4
                and    ccm.cost_cmpntcls_code = '3_감가상각'
                group by pde.organization_id
                     , ccm.cost_cmpntcls_id  ) 
    """
    global verify_param
    verify_param = (org, period, cost_type, calendar)
    cursor.execute(verifysql, verify_param)
    verifydata = cursor.fetchall()
    #print(verifydata)
    # verifydata = verifydata.fillna(0)
    print(len(verifydata))
    direct_labor_amt = 0
    direct_machine_amt = 0
    var_consume_amt = 0
    fixed_labor_amt = 0
    mfg_expense_amt = 0
    emp_benefit_amt = 0
    depreciation_amt = 0
    for idx in range(len(verifydata)-2):
        if verifydata[idx][17] is None:
            direct_labor_amt = direct_labor_amt + 0
        else:
            direct_labor_amt = direct_labor_amt + verifydata[idx][17]
        if verifydata[idx][18] is None:
            direct_machine_amt = direct_machine_amt + 0
        else:
            direct_machine_amt = direct_machine_amt + verifydata[idx][18]
        if verifydata[idx][19] is None:
            var_consume_amt = var_consume_amt + 0
        else:
            var_consume_amt = var_consume_amt + verifydata[idx][19]
        if verifydata[idx][20] is None:
            fixed_labor_amt = fixed_labor_amt + 0
        else:
            fixed_labor_amt = fixed_labor_amt + verifydata[idx][20]
        if verifydata[idx][21] is None:
            mfg_expense_amt = mfg_expense_amt + 0
        else:
            mfg_expense_amt = mfg_expense_amt + verifydata[idx][21]
        if verifydata[idx][22] is None:
            emp_benefit_amt = emp_benefit_amt + 0
        else:
            emp_benefit_amt = emp_benefit_amt + verifydata[idx][22]
        if verifydata[idx][23] is None:
            depreciation_amt = depreciation_amt + 0
        else:
            depreciation_amt = depreciation_amt + verifydata[idx][23]
    # print(direct_labor_amt)
    # print(direct_machine_amt)
    # print(var_consume_amt)
    # print(fixed_labor_amt)
    # print(mfg_expense_amt)
    # print(emp_benefit_amt)
    # print(depreciation_amt)
    exp_direct_labor_amt    = verifydata[len(verifydata)-1][17]
    exp_direct_machine_amt  = verifydata[len(verifydata)-1][18]
    exp_var_consume_amt     = verifydata[len(verifydata)-1][19]
    exp_fixed_labor_amt     = verifydata[len(verifydata)-1][20]
    exp_mfg_expense_amt     = verifydata[len(verifydata)-1][21]
    exp_emp_benefit_amt     = verifydata[len(verifydata)-1][22]
    exp_depreciation_amt    = verifydata[len(verifydata)-1][23]
    # print(exp_direct_labor_amt)
    cursor.close()
    return render_template('verify.html'
                           , org_list=org_list
                           , cost_type_list=cost_type_list
                           , period_list=period_list
                           , calendar_list=calendar_list
                           , data_list=verifydata
                           , direct_labor_amt=direct_labor_amt
                           , direct_machine_amt=direct_machine_amt
                           , var_consume_amt=var_consume_amt
                           , fixed_labor_amt=fixed_labor_amt
                           , mfg_expense_amt=mfg_expense_amt
                           , emp_benefit_amt=emp_benefit_amt
                           , depreciation_amt=depreciation_amt
                           , exp_direct_labor_amt=exp_direct_labor_amt
                           , exp_direct_machine_amt=exp_direct_machine_amt
                           , exp_var_consume_amt=exp_var_consume_amt
                           , exp_fixed_labor_amt=exp_fixed_labor_amt
                           , exp_mfg_expense_amt=exp_mfg_expense_amt
                           , exp_emp_benefit_amt=exp_emp_benefit_amt
                           , exp_depreciation_amt=exp_depreciation_amt
                           )

@app.route('/verify_std', methods = ['GET','POST'])
# def extract(org=None, period=None, cost_type=None):
def verify_std():
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
    # conn = cx_Oracle.connect(cfg.username, cfg.password, cfg.dsn, encoding=cfg.encoding)
    # cursor = conn.cursor()
    # print(cursor)
    global verifysql
    verifysql="""
        select decode(org, null, '합계', org) org
     , item
     , description
     , gl_cls
     , inventory_item_id
     , quantity1 "계획수량"
--     , ind_qty
     , std_cost "표준원가"
     , c1 "원재료비"
     , c2 "포장재비"
     , c3 "직접노무"
     , c4 "기계가동"
     , c5 "변동소모"
     , c6 "고정노무"
     , c7 "생산경비"
     , c8 "복리후생"
     , c9 "감가상각"
     , sum(c3s) "직접노무금액"
     , sum(c4s) "기계가동금액"
     , sum(c5s) "변동소모금액"
     , sum(c6s) "고정노무금액"
     , sum(c7s) "생산경비금액"
     , sum(c8s) "복리후생금액"
     , sum(c9s) "감가상각금액"
from   (select org
             , item
             , description
             , gl_cls
             , inventory_item_id
             , quantity1
--             , ind_qty
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
--                     , pmq.quantity1 ind_qty
                     , decode(ccd.cost_cmpntcls_id, 1, sum(ccd.cmpnt_cost)  , null) as c1
                     , decode(ccd.cost_cmpntcls_id, 2, sum(ccd.cmpnt_cost)  , null) as c2
                     , decode(ccd.cost_cmpntcls_id, 3, sum(ccd.cmpnt_cost)  , null) as c3
                     , decode(ccd.cost_cmpntcls_id, 4, sum(ccd.cmpnt_cost)  , null) as c4
                     , decode(ccd.cost_cmpntcls_id, 5, sum(ccd.cmpnt_cost)     , null) as c5
                     , decode(ccd.cost_cmpntcls_id, 6, sum(ccd.cmpnt_cost)     , null) as c6
                     , decode(ccd.cost_cmpntcls_id, 7, sum(ccd.cmpnt_cost)     , null) as c7
                     , decode(ccd.cost_cmpntcls_id, 8, sum(ccd.cmpnt_cost)     , null) as c8
                     , decode(ccd.cost_cmpntcls_id, 9, sum(ccd.cmpnt_cost)     , null) as c9
                     , decode(ccd.cost_cmpntcls_id, 3, sum(ccd.cmpnt_cost)  , null)*xpq.quantity1 as c3s
                     , decode(ccd.cost_cmpntcls_id, 4, sum(ccd.cmpnt_cost)  , null)*xpq.quantity1 as c4s
                     , decode(ccd.cost_cmpntcls_id, 5, sum(ccd.cmpnt_cost)     , null)*xpq.quantity1 as c5s
                     , decode(ccd.cost_cmpntcls_id, 6, sum(ccd.cmpnt_cost)     , null)*xpq.quantity1 as c6s
                     , decode(ccd.cost_cmpntcls_id, 7, sum(ccd.cmpnt_cost)     , null)*xpq.quantity1 as c7s
                     , decode(ccd.cost_cmpntcls_id, 8, sum(ccd.cmpnt_cost)     , null)*xpq.quantity1 as c8s
                     , decode(ccd.cost_cmpntcls_id, 9, sum(ccd.cmpnt_cost)     , null)*xpq.quantity1 as c9s
                from   cm_cmpt_dtl ccd
                     , mtl_system_items_b msi
                     , gmf_organization_definitions god
                     , mtl_categories mc
                     , mtl_item_categories mic
                     , xcstf_plan_order_quantities xpq
--                     , xcstf_business_plan_mfg_qty_v pmq
--                     , cm_brdn_dtl cbd
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
--                and    ccd.period_id = pmq.period_id(+)
--                and    ccd.cost_type_id = pmq.cost_type_id(+)
                and    mc.segment1 in ('05','06')
                and    god.organization_id = xpq.organization_id(+)
                and    msi.inventory_item_id = xpq.inventory_item_id(+)
                and    ccd.period_id = xpq.period_id(+)
--                and    god.organization_id = pmq.organization_id(+)
--                and    msi.inventory_item_id = pmq.inventory_item_id(+)
--                and    pmq.organization_id = cbd.organization_id(+)
--                and    pmq.inventory_item_id = cbd.inventory_item_id(+)
--                and    pmq.period_id = cbd.period_id(+)
--                and    pmq.cost_type_id = cbd.cost_type_id(+)
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
--                     , pmq.quantity1
                     , ccd.cost_cmpntcls_id
--                     , cbd.cost_cmpntcls_id
--                     , cbd.burden_usage 
                     )
        group by org
             , item
             , description
             , gl_cls
             , inventory_item_id
             , quantity1
--             , ind_qty 
             )
group by rollup(( org
                     , item
                     , description
                     , gl_cls
                     , inventory_item_id
                     , quantity1
--                     , ind_qty
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
--     , null
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
             , pev.cost_cmpntcls_id
        union all
        select pde.organization_id, ccm.cost_cmpntcls_id, sum(pde.expense_amount) expense_amount 
        from xcstf_plan_dep_expenses pde
        , gmf_organization_definitions god
        , gmf_period_statuses gps
        , cm_mthd_mst cmm
        , cm_cmpt_mst ccm
        where  1=1
        and    pde.organization_id = god.organization_id
        and    pde.period_id = gps.period_id
        and    cmm.cost_type_id = pde.cost_type_id
        and    pde.cost_type_id = gps.cost_type_id
        and    god.organization_code = :1
        and    gps.period_code = :2
        and    cmm.cost_mthd_code = :3
        and    gps.calendar_code = :4
        and    ccm.cost_cmpntcls_code = '3_감가상각'
        group by pde.organization_id
             , ccm.cost_cmpntcls_id     
             ) 
    """
    global verify_param
    verify_param = (org, period, cost_type, calendar)
    cursor.execute(verifysql, verify_param)
    verifydata = cursor.fetchall()
    #print(verifydata)
    # verifydata = verifydata.fillna(0)
    print(len(verifydata))
    direct_labor_amt = 0
    direct_machine_amt = 0
    var_consume_amt = 0
    fixed_labor_amt = 0
    mfg_expense_amt = 0
    emp_benefit_amt = 0
    depreciation_amt = 0
    for idx in range(len(verifydata)-2):
        if verifydata[idx][16] is None:
            direct_labor_amt = direct_labor_amt + 0
        else:
            direct_labor_amt = direct_labor_amt + verifydata[idx][16]
        if verifydata[idx][17] is None:
            direct_machine_amt = direct_machine_amt + 0
        else:
            direct_machine_amt = direct_machine_amt + verifydata[idx][17]
        if verifydata[idx][18] is None:
            var_consume_amt = var_consume_amt + 0
        else:
            var_consume_amt = var_consume_amt + verifydata[idx][18]
        if verifydata[idx][19] is None:
            fixed_labor_amt = fixed_labor_amt + 0
        else:
            fixed_labor_amt = fixed_labor_amt + verifydata[idx][19]
        if verifydata[idx][20] is None:
            mfg_expense_amt = mfg_expense_amt + 0
        else:
            mfg_expense_amt = mfg_expense_amt + verifydata[idx][20]
        if verifydata[idx][21] is None:
            emp_benefit_amt = emp_benefit_amt + 0
        else:
            emp_benefit_amt = emp_benefit_amt + verifydata[idx][21]
        if verifydata[idx][22] is None:
            depreciation_amt = depreciation_amt + 0
        else:
            depreciation_amt = depreciation_amt + verifydata[idx][22]
    # print(direct_labor_amt)
    # print(direct_machine_amt)
    # print(var_consume_amt)
    # print(fixed_labor_amt)
    # print(mfg_expense_amt)
    # print(emp_benefit_amt)
    # print(depreciation_amt)
    exp_direct_labor_amt    = verifydata[len(verifydata)-1][16]
    exp_direct_machine_amt  = verifydata[len(verifydata)-1][17]
    exp_var_consume_amt     = verifydata[len(verifydata)-1][18]
    exp_fixed_labor_amt     = verifydata[len(verifydata)-1][19]
    exp_mfg_expense_amt     = verifydata[len(verifydata)-1][20]
    exp_emp_benefit_amt     = verifydata[len(verifydata)-1][21]
    exp_depreciation_amt    = verifydata[len(verifydata)-1][22]
    # print(exp_direct_labor_amt)
    cursor.close()
    return render_template('verify_std.html'
                           , org_list=org_list
                           , cost_type_list=cost_type_list
                           , period_list=period_list
                           , calendar_list=calendar_list
                           , data_list=verifydata
                           , direct_labor_amt=direct_labor_amt
                           , direct_machine_amt=direct_machine_amt
                           , var_consume_amt=var_consume_amt
                           , fixed_labor_amt=fixed_labor_amt
                           , mfg_expense_amt=mfg_expense_amt
                           , emp_benefit_amt=emp_benefit_amt
                           , depreciation_amt=depreciation_amt
                           , exp_direct_labor_amt=exp_direct_labor_amt
                           , exp_direct_machine_amt=exp_direct_machine_amt
                           , exp_var_consume_amt=exp_var_consume_amt
                           , exp_fixed_labor_amt=exp_fixed_labor_amt
                           , exp_mfg_expense_amt=exp_mfg_expense_amt
                           , exp_emp_benefit_amt=exp_emp_benefit_amt
                           , exp_depreciation_amt=exp_depreciation_amt
                           )

if __name__ == '__main__':
    #서버 실행
    #app.run(debug = True)
    app.run(host='0.0.0.0', port=5000)