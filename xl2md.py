#!/usr/bin/env python

# %%
# openpyxl requied
import pandas as pd

# %%
# use the second row as the column name
# ignore the third and forth row
# pandas's read_excel is based on openpyxl
# the input xlsx is specified by command line argument

import sys
if len(sys.argv) > 1:
    df = pd.read_excel(sys.argv[1], header=1, skiprows=[2,3])
else:
    # print error message
    print("Usage: python xl2md.py <input xlsx> <output md>")
    print("Example: python xl2md.py 12.17音乐会节目信息征集.xlsx program.md")
    # exit the program
    sys.exit(1)




# %%
# define a markdown front matter
front_matter="""---
marp: true
theme: uncover
class: invert
---

<style>
  table {
    font-size: 30px;
  }
  th,td {
    border: none!important;
  }
</style>
"""

# create a new markdown file named program prensentation which begins with the `front_matter`
# write to the file specified by the second command line argument
with open(sys.argv[2], 'w') as f:
    f.write(front_matter)



# %%
# the first page of the program presentation 
frontpage = """# 时光流转 琴音不辍
浙江大学研究生艺术团钢琴音乐会
"""

# write the frontpage
with open(sys.argv[2], 'a') as f:
    f.write(frontpage)
    f.write('\n---\n')

# the second page is warping
warping = """# 注意事项
1. 请关闭您的手机或将您的手机调至静音状态
1. 请勿使用闪光灯拍照、摄影
1. 请妥善放置塑料袋等容易发出声响的物品
1. 在演出期间，请您尽量避免在场内来回走动
"""

# write the warping page
with open(sys.argv[2], 'a') as f:
    f.write(warping)
    f.write('\n---\n')

# %%
# for each row in df, conver it and write to the markdown file as follows
# 曲目中文名 --> ### 曲目中文名
# 曲目英文名 --> **曲目英文名**
# 作曲家中文 作曲家英文 --> ##### 作曲家中文 作曲家英文
# 表演者 表演者拼音 --->
# |       |      |
# | :-----|:------|
# | 表演者 | 表演者 |
# | Performer | 表演者拼音 |
# and end with --- except the last row

with open(sys.argv[2], 'a') as f:
    for index, row in df.iterrows():

        if row["曲目中文名"] == "中场休息":
            
            f.write(f'# {row["曲目中文名"]}\n')
            f.write(f'10分钟\n')


        else:
            f.write(f'### {row["曲目中文名"]}\n')
            f.write(f'**{row["曲目英文名"]}**\n')
            f.write(f'##### {row["作曲家中文"]} {row["作曲家英文"]}\n')
            f.write(f'|       |      |\n')
            f.write(f'| :-----|:------|\n')
            f.write(f'| 表演者 | {row["表演者"]} |\n')
            f.write(f'| Performer | {row["表演者拼音"]} |\n')

        f.write('\n---\n')


# %%
# the last page is the same as the front page
with open(sys.argv[2], 'a') as f:
    f.write(frontpage)

# print out successful message
print(f"Successfully convert the {sys.argv[1]} to {sys.argv[2]}")