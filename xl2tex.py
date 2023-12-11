#!/usr/bin/env python
import os.path
import pandas as pd
from load_xl import load_xl

df = load_xl(purpose="tex")


# for each row in df, convert it and write to a tex file called pieces.tex as follows
# \piece{曲目中文名}{曲目英文名}{作曲家中文}{作曲家英文}{表演者}{表演者拼音}{改编者}

intermission="""\\noindent\\begin{center}\\normalsize \hspace{-1.5cm} 中场休息 Intermission\\end{center}
\\vspace{0.3cm}
"""

with open("pieces.tex", 'w', encoding='utf-8') as f:
    running_mvt_name = ""
    for index, row in df.iterrows():

        # compare the current row with the next row
        if index < len(df) - 1:
            next_row = df.iloc[index + 1]
            if row["曲目中文名"] == next_row["曲目中文名"]:
                running_mvt_name += str(row["乐章英文 (选填) "])
                running_mvt_name += " "
                print(running_mvt_name)
                continue
            else:
                if row["乐章英文 (选填) "] == row["乐章英文 (选填) "]: 
                    if row["乐章英文 (选填) "] != '\u3000':
                        running_mvt_name += str(row["乐章英文 (选填) "])

        if row["曲目中文名"] == "中场休息":
            f.write(intermission)
        else:

            # replace all \n to \\ in 表演者中文
            row["表演者中文"] = row["表演者中文"].replace("\n", "\\\\")


            f.write(f'\\piece{{{row["曲目中文名"]}}}{{{row["曲目英文名"]}}}{{{row["作曲家中文"]}}}{{{row["作曲家英文"]}}}{{{row["表演者中文"]}}}')
            
            # if 表演者英文 is NAN
            if row["表演者英文"] == row["表演者英文"]: 
                f.write(f'{{{row["表演者英文"]}}}')
            else:
                f.write('{}')
            if row["改编者 (选填) "] == row["改编者 (选填) "]: 
                if row["改编者 (选填) "] != '\u3000':
                    f.write(f'{{{row["改编者 (选填) "]}}}')
                else:
                    print
                    f.write('{}')
            else:
                f.write('{}')
            
            if running_mvt_name != "":
                f.write(f'{{{running_mvt_name}}}')
                running_mvt_name = ""
            elif row["乐章英文 (选填) "] == row["乐章英文 (选填) "]: 
                if row["乐章英文 (选填) "] != '\u3000':
                    f.write(f'{{{row["乐章英文 (选填) "]}}}')
                else:
                    print("there")
                    f.write('{}')
            else:
                f.write('{}')
            
            f.write('\n\n')


    print("File pieces.tex generated successfully!")