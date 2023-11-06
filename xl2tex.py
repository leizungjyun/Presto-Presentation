#!/usr/bin/env python
import os.path
import pandas as pd

# %%
# use the second row as the column name
# ignore the third and forth row
# pandas's read_excel is based on openpyxl
import sys
if len(sys.argv) > 1:
    # decide if file exists
    if not os.path.isfile(sys.argv[1]):
        print(f"File {sys.argv[1]} does not exist!")
        sys.exit(1)
    else:
        df = pd.read_excel(sys.argv[1], header=1, skiprows=[2,3])
else:
    # print error message
    print("Usage: python xl2tex.py <input xlsx>")
    print("Example: python xl2md.py 12.17音乐会节目信息征集.xlsx")
    # exit the program
    sys.exit(1)

# for each row in df, convert it and write to a tex file called pieces.tex as follows
# \piece{曲目中文名}{曲目英文名}{作曲家中文}{作曲家英文}{表演者}{表演者拼音}{改编者}

intermission="""\\noindent\\begin{center}\\normalsize \hspace{-1.5cm} 中场休息 Intermission\\end{center}
\\vspace{0.3cm}
"""

with open("pieces.tex", 'w') as f:
    for index, row in df.iterrows():
        if row["曲目中文名"] == "中场休息":
            f.write(intermission)
        else:
            # if row["改编者"] is NAN
            if row["改编者"] == row["改编者"]:
                f.write(f'\\piece{{{row["曲目中文名"]}}}{{{row["曲目英文名"]}}}{{{row["作曲家中文"]}}}{{{row["作曲家英文"]}}}{{{row["表演者"]}}}{{{row["表演者拼音"]}}}{{{row["改编者"]}}}\n\n')
            else:
                f.write(f'\\piece{{{row["曲目中文名"]}}}{{{row["曲目英文名"]}}}{{{row["作曲家中文"]}}}{{{row["作曲家英文"]}}}{{{row["表演者"]}}}{{{row["表演者拼音"]}}}{{}}\n\n')
    print("File pieces.tex generated successfully!")