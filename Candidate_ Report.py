import pandas as pd
import smartsheet
import win32com.client
import numpy as np
import time
import schedule
from smartsheet.models import Report
from smartsheet_dataframe import get_as_df, get_report_as_df, get_sheet_as_df
from pretty_html_table import build_table


smartsheet_client = smartsheet.Smartsheet("3v5MObKahNM8NEhDeicvHFm6uLwzy8LQhYANN")
rec_rep = smartsheet_client.Sheets.get_sheet("6448692810698628")
df = get_sheet_as_df(sheet_obj=rec_rep)

can_rep1 = df[['Condition','Position','Candidate Name','Status','Last Correspondence','Next Steps','Phone Screen Date','Screener','Panel Interview Date','Panel Interviewers']].copy()
can_rep2 = can_rep1.query("(Condition =='Moving Forward') or (Condition =='Under Review') or (Condition =='Hired')")


outlook = win32com.client.Dispatch("outlook.application")
mail = outlook.CreateItem(0)
mail.To = 'tmenachof@primepartnersengineering.com'
mail.Subject = "Candidate Report"

Table = build_table(can_rep2, 'blue_dark',
                    font_size = '10px',
                    text_align = 'center',
                    width_dict = ['150px','150px','150px','250px','auto','200px','100px','auto','100px','150px'],
                    )

html = f"""
<html><body><p style="font-family:garamond">Hi all,
<br>
<br>Please find the Current Candidate report below. Let me know if you have any questions.</p>
<br>{Table}
<p style="font-family:garamond">Best,
<br>
<br>Tristan</p>
</body></html>
"""
mail.HTMLBody = html
mail.Send()
