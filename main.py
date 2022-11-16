import win32com.client
import os
import pandas as pd

path = os.getcwdb()
file_list = [file.decode('utf8') for file in os.listdir(path)]
file_list_cleaned = [file for file in file_list if file.endswith('.pptx') and not file.startswith('~$')]
print(file_list_cleaned)
df = pd.DataFrame()


for i, ppt in enumerate(file_list_cleaned):
    ppt_app = win32com.client.GetObject(ppt)
    df.loc['00_File'] = ppt
    print("----Ppt name----",ppt)
    for j, ppt_slide in enumerate(ppt_app.Slides):
        df.loc[i,'00_Slide'] = j
        print("----Slide number----", ppt_slide, '--#--', j)
        for k, comment in enumerate(ppt_slide.Comments):
            print("----Comment----", ppt_slide, '--#--', k)
            df.loc[i, str(k)+"_Comment" ] = comment.Text
            for l, reply in enumerate(comment.Replies):
                print("----Reply---", ppt_slide, '--#--', l)
                df.loc[i,str(l)+"_Replies"] = reply.Text


df = df.reindex(sorted(df.columns), axis=1)
df.to_excel(excel_writer = 'comments.xlsx')
print(df)


