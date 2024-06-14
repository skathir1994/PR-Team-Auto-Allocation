# Importing the necessary packages
import pandas as pd
import datetime
import openpyxl as open1
import win32com.client
import mysql.connector
#
# Reading roster base file
#
roster = r'D:\Users\skathir\Desktop\PR Roster\Roster.xlsx'
data = pd.read_excel(roster)
#
# Creating mysql connection
#
con = mysql.connector.connect(
    host="cmtlabs-db-pr.aka.amazon.com",
    user="pr_reader",
    password="************",
    database="price_rejects",
)
#
# Creating new class for getting current HC from roster base.
#
class allocation:
    def roster(self):
        global condition1
        column_headers = data.columns
        for column_header in column_headers:  # featching the today rosted ID's
            try:
                if column_header.date() == datetime.date.today():
                    result = data[column_header]
            except AttributeError:
                pass

        df = pd.DataFrame(
            data=data['ID'],
        )
        df1 = pd.DataFrame(
            data=result,
        )
        df_total = [df, df1]
        full = pd.concat(
            df_total,
            axis=1,
        )
        full.columns = ['Auditor_Id', 'Shift']
        df3 = pd.DataFrame(
            data=full,
        )
        condition = df3[df3['Shift'] == 'GS']
        condition1 = condition['Auditor_Id']
        return
#
# Creating new SQL class for getting current pending listings from mysql.
#
    def sql_query(self):
        res = con.cursor()
        sql = "select region,merchant_id,asin,competitor,reject_code,recommended_price,reject_date,buyer_login,gl,category_code,competitor_url,competitor_price,competitorshipping_price,uploaded_date,status,super_status from price_rejects.price_reject_new where status='allocated' and region not in (6,44571) ORDER BY reject_code;"
        res.execute(sql)
        result = res.fetchall()
        df_base = pd.DataFrame(
            data=result,
            columns=[
                "region",
                "merchant_id",
                "asin",
                "competitor",
                "reject_code",
                "recommended_price",
                "reject_date",
                "buyer_login",
                "gl",
                "categoryCode",
                "competitor_url",
                'competitor_price',
                'competitorshipping_price',
                'uploaded_date',
                'status',
                'super_status',
            ],
        )
        df_base['Rejected_Date'] = df_base['reject_date'].dt.strftime('%m-%d-%Y')
        df_base['reject_code1'] = df_base['reject_code']
        df_base1 = df_base.sort_values(
            "reject_code1",
            ascending=True,
        )
        df_final = pd.DataFrame(df_base1)
        total = pd.DataFrame(data=[], columns=['Auditor_Id'])
        lst = []
        total_id = condition1
        i = 0
        for i in range(len(df_final)):
            for j in total_id:
                s = j
                lst.append(s)
                # print(s)
        total['Auditor_Id'] = lst
        df_base2 = pd.concat([df_final, total], axis=1, join='inner', sort=False)
        df_pivot = pd.pivot_table(
            df_base2,
            index=['Auditor_Id', 'reject_code'],
            columns=['Rejected_Date'],
            values='asin',
            aggfunc='count',
            margins=True,
            fill_value=" ",
        )

        print(df_pivot)

        table_html = df_pivot.to_html(table_id='pending_no')
#
# SQL query for getting IN & JP mkpl pending listings.
#
        sql1 = "select region,merchant_id,asin,competitor,reject_code,recommended_price,reject_date,buyer_login,gl,category_code,competitor_url,competitor_price,competitorshipping_price,uploaded_date,status,super_status from price_rejects.price_reject_new where status='allocated' and region in (6,44571);"
        res.execute(sql1)
        result1 = res.fetchall()
        df_jp = pd.DataFrame(
            data=result1,
            columns=[
                "region",
                "merchant_id",
                "asin",
                "competitor",
                "reject_code",
                "recommended_price",
                "reject_date",
                "buyer_login",
                "gl",
                "categoryCode",
                "competitor_url",
                'competitor_price',
                'competitorshipping_price',
                'uploaded_date',
                'status',
                'super_status',
            ],
        )
        df_jp['Rejected_Date'] = df_jp['reject_date'].dt.strftime('%m-%d-%Y')

        total1 = pd.DataFrame(data=[], columns=['Auditor_Id'])

        df_total2 = pd.concat(
            [df_jp, total1],
            axis=1,
        )

        df_total2.loc[df_total2['region'] == '6', 'Auditor_Id'] = 'kashok'
        df_total2.loc[df_total2['region'] == '44571', 'Auditor_Id'] = 'hgoyal'
#
# Creating pivot.
#
        df_pivot1 = pd.pivot_table(
            df_total2,
            index=['Auditor_Id', 'reject_code'],
            columns=['Rejected_Date'],
            values='asin',
            aggfunc='count',
            margins=True,
            fill_value=" ",
        )
#
# Pivot to HTML convert.
#
        table_html1 = df_pivot1.to_html(table_id='pending_no')
#
# Read the relevant text file.
#
        with open(
            r'D:\Users\skathir\PycharmProjects\PR_Daily_Allocation\table style.txt'
        ) as file:
            table_style = file.read()

        with open(
            r'D:\Users\skathir\PycharmProjects\PR_Daily_Allocation\before_table.txt'
        ) as body_file:
            body_html = body_file.read()

        with open(
            r'D:\Users\skathir\PycharmProjects\PR_Daily_Allocation\after_table.txt'
        ) as last_file:
            last_html = last_file.read()

        final_html = (
            body_html
            + table_html
            + '<br/> IN & JP MKPL <br/>'
            + table_html1
            + last_html
        )
#
# Get the current date.
#
        df2 = pd.to_datetime('today').strftime('%m-%d-%Y')
#
# Outlook creation.
#
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = 'pricerejects-team@amazon.com'
        mail.CC = 'pr-managers@amazon.com'
        mail.Subject = 'PR Work Allocation on' + ' ' + df2
        mail.HTMLBody = final_html
        mail.Send()
        return

#
# Function call.
#
Object = allocation()
Object.roster()
Object.sql_query()
