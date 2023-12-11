#!/usr/bin/env python
import pandas as pd # openpyxl requied
import sys
from load_xl import load_xl


# load the excel file and return dataframe, heading_cn, heading_en, troupe_cn, troupe_en
df, heading_cn, heading_en, troupe_cn, troupe_en = load_xl(purpose="slides")

frontpage = f"# {heading_cn}\n#### {heading_en}\n# \n#### {troupe_cn}\n###### {troupe_en}\n"

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
h1,h3,h4{
  margin-bottom:00px;
}
</style>
"""
# the second page is warning
warning = """# 注意事项
1. 请关闭您的手机或将您的手机调至静音状态
1. 请勿使用闪光灯拍照、摄影
1. 请妥善放置塑料袋等容易发出声响的物品
1. 在演出期间，请您尽量避免在场内来回走动
"""

movement_warning = """
此作品有多个乐章，为保证演出效果，请勿在乐章间鼓掌。
This work has multiple movements.
Please do not applaud between movements.
"""


with open(md_file, 'w', encoding='utf-8') as f:
    f.write(front_matter)
    f.write(frontpage)
    f.write('\n---\n')
    f.write(warning)
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


    last_row = None
    for index, row in df.iterrows():
        
        if last_row is not None:
            if row["曲目中文名"] == last_row["曲目中文名"]:
                f.write(f'{movement_warning}\n')
                f.write('\n---\n')
        if row["曲目中文名"] == "中场休息":
            f.write(f'# {row["曲目中文名"]}\n')
            f.write(f'### Intermission\n')
            f.write(f'### 10分钟\n')
        else:
            f.write(f'##### {row["作曲家中文"]} {row["作曲家英文"]}\n')
            if row["改编者 (选填) "] == row["改编者 (选填) "]:
                if row["改编者 (选填) "] !='\u3000':
                    f.write(f'###### {row["改编者 (选填) "]} 改编\n')
            f.write(f'### {row["曲目中文名"]}\n')
            f.write(f'**{row["曲目英文名"]}**\n\n')
            if row["乐章英文 (选填) "] == row["乐章英文 (选填) "]: 
                if row["乐章英文 (选填) "] !='\u3000':
                    f.write(f'**{row["乐章英文 (选填) "]}**\n')
            # if 改编者 is not nan
            f.write(f'|         |\n')
            f.write(f'| :-----: |\n')
            f.write(f'|         |\n')
            f.write(f'|  {row["表演者中文"]}  |\n')

            # if 表演者英文 is not NAN
            if row["表演者英文"] == row["表演者英文"]: 
                if row["表演者英文"] !='\u3000':
                    f.write(f'|  {row["表演者英文"]} |\n')

            # f.write(f'| Performer | {row["表演者英文"]} |\n')
        f.write('\n---\n')
        last_row = row


    # the last page is the same as the front page
    f.write(frontpage)

# print out successful message
print(f"Successfully convert the {sys.argv[1]} to {md_file}")