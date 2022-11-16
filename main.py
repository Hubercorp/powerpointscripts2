import win32com.client
import os
import pandas as pd

path = os.getcwdb()
file_list = [file.decode('utf8') for file in os.listdir(path)]
file_list_cleaned = [file for file in file_list if file.endswith('.pptx') and not file.startswith('~$')]
print(file_list_cleaned)
df = pd.DataFrame()


for ppt in file_list_cleaned:
    ppt_app = win32com.client.GetObject(path.decode('utf8') +"\\" + ppt)
    for j, ppt_slide in enumerate(ppt_app.Slides):
        for k, comment in enumerate(ppt_slide.Comments):
            if comment is not None:
                df.loc[j,'00_File'] = ppt
                df.loc[j,'00_Slide'] = j
                df.loc[j, str(k)+"_Comment" ] = comment.Text
                for l, reply in enumerate(comment.Replies):
                    df.loc[j,str(l)+"_Replies"] = reply.Text


df = df.reindex(sorted(df.columns), axis=1)
df.to_excel(excel_writer = 'comments.xlsx')
print(df)





