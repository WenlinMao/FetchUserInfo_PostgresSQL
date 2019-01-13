import psycopg2
import psycopg2.extras
import xlsxwriter
import datetime
from DBConfig import *

try:
    # production db
    db = psycopg2.connect(PRODUCTION_URL)

    # development db
    #db = psycopg2.connect(DEVELOPMENT_URL)

    cur = db.cursor(cursor_factory=psycopg2.extras.DictCursor)

    # take first_name, last_name, user_id from user_profiles
    cur.execute("select distinct on (first_name, last_name) * from user_profiles")
    rows = cur.fetchall()

    # create xlsx using xlsxwriter
    with xlsxwriter.Workbook('User_email.xlsx') as workbook:
        worksheet = workbook.add_worksheet()

        # row and col to write in xlsx
        row = 0
        
        worksheet.write(row, 0, 'First Name')
        worksheet.write(row, 1, 'Last Name')
        worksheet.write(row, 2, 'Email')
        worksheet.write(row, 3, 'Instituion')
        worksheet.write(row, 4, 'Primary Stakeholder')
        worksheet.write(row, 5, 'Secondary Stakeholder')
        worksheet.write(row, 6, 'Created At')
        worksheet.write(row, 7, 'Updated At')

        for elt in rows:
            row += 1
            
            user_id = elt["user_id"]
            SQL = "SELECT DISTINCT email FROM users WHERE id = %s"
            cur.execute(SQL, (user_id,))

            email = cur.fetchone()[0]
            first_name = elt["first_name"]
            last_name = elt["last_name"]
            institution = elt["institution"]
            pri_stakeholder = elt["primary_stakeholder_group"]

            # list of groups, convert list to str
            sec_stakeholder = ''.join(elt["secondary_stakeholder_groups"])

            # convert datetime object to string
            created_date = elt["created_at"].strftime('%m/%d/%Y')
            updated_date = elt["updated_at"].strftime('%m/%d/%Y')

            worksheet.write(row, 0, first_name)
            worksheet.write(row, 1, last_name)
            worksheet.write(row, 2, email)
            worksheet.write(row, 3, institution)
            worksheet.write(row, 4, pri_stakeholder)
            worksheet.write(row, 5, sec_stakeholder)
            worksheet.write(row, 6, created_date)
            worksheet.write(row, 7, updated_date)

except:
    raise




