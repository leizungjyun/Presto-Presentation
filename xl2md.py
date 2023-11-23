#!/usr/bin/env python
import pandas as pd # openpyxl requied
import os.path
import sys

# use the second row as the column name
# pandas's read_excel is based on openpyxl
# the input xlsx is specified by command line argument
if len(sys.argv) > 1:
    # decide if file exists
    if not os.path.isfile(sys.argv[1]):
        print(f"File {sys.argv[1]} does not exist!")
        sys.exit(1)
    else:
# ignore the third to the nineth rows
        df = pd.read_excel(sys.argv[1], header=1, skiprows=range(2, 9))
else:
    # print error message
    print("Usage: python xl2md.py <input xlsx>")
    print("Example: python xl2md.py 12.17音乐会节目信息征集.xlsx")
    sys.exit(1)

# create a new markdown file named program prensentation which begins with the `front_matter`
# write to the file specified by the second command line argument
md_file = "slides.md"

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
section {
    font-family: "Garamond";
}
h3{
  margin-bottom:00px;
}
</style>
"""
# the second page is warping
warping = """# 注意事项
1. 请关闭您的手机或将您的手机调至静音状态
1. 请勿使用闪光灯拍照、摄影
1. 请妥善放置塑料袋等容易发出声响的物品
1. 在演出期间，请您尽量避免在场内来回走动
"""
# the first page of the program presentation 
frontpage = """# 冬季学术交流音乐会
中国音乐学院管弦系 浙江大学研究生艺术团
"""

with open(md_file, 'w', encoding='utf-8') as f:
    f.write(front_matter)
    f.write(frontpage)
    f.write('\n---\n')
    f.write(warping)
    f.write('\n---\n')
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
    for index, row in df.iterrows():
        if row["曲目中文名"] == "中场休息":
            f.write(f'# {row["曲目中文名"]}\n')
            f.write(f'10分钟\n')
        else:
            f.write(f'##### {row["作曲家中文"]} {row["作曲家英文"]}\n')
            f.write(f'### {row["曲目中文名"]}\n')
            f.write(f'**{row["曲目英文名"]}**\n\n')
            if row["乐章英文 (选填) "] == row["乐章英文 (选填) "]: 
                f.write(f'**{row["乐章英文 (选填) "]}**\n')
            # if 改编者 is not nan
            if row["改编者 (选填) "] == row["改编者 (选填) "]: 
                f.write(f'###### {row["改编者 (选填) "]} 改编\n')
            f.write(f'|       |      |\n')
            f.write(f'| :-----|:------|\n')
            f.write(f'|       |      |\n')
            f.write(f'|  {row["表演者中文"]} |\n')

            # if 表演者英文 is not NAN
            if row["表演者英文"] == row["表演者英文"]: 
                f.write(f'|  {row["表演者英文"]} |\n')

            # f.write(f'| Performer | {row["表演者英文"]} |\n')
        f.write('\n---\n')

    # the last page is the same as the front page
    f.write(frontpage)

# print out successful message
print(f"Successfully convert the {sys.argv[1]} to {md_file}")