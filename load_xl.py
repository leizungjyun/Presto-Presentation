#!/usr/bin/env python
import pandas as pd # openpyxl requied
import os.path
import sys

def load_xl(purpose):
    # use the sixth row as the column name
    # pandas's read_excel is based on openpyxl
    # the input xlsx is specified by command line argument
    if len(sys.argv) > 1:
        # decide if file exists
        if not os.path.isfile(sys.argv[1]):
            print(f"File {sys.argv[1]} does not exist!")
            sys.exit(1)
        else:
            # ignore the 8th to the 13th rows
            df = pd.read_excel(sys.argv[1], header=6, skiprows=range(7, 14))

            if purpose == 'slides':
                # heading_cn in b1
                # heading_en in b2
                # troupe_cn in b3
                # troupe_en in b4
                # load the file, don't skip the heading
                head_df = pd.read_excel(sys.argv[1], header=None)

                # load heading_cn, heading_en, troupe_cn, troupe_en
                heading_cn = head_df.iloc[0, 1]
                heading_en = head_df.iloc[1, 1]
                troupe_cn = head_df.iloc[2, 1]
                troupe_en = head_df.iloc[3, 1]

                return df, heading_cn, heading_en, troupe_cn, troupe_en

            elif purpose == 'tex':
                return df
            else:
                raise ValueError("purpose must be either presentation or tex")
    else:
        # print error message
        if purpose == "presentation":
            print("Usage: python xl2md.py <input xlsx>")
            print("Example: python xl2md.py 12.17音乐会节目信息征集.xlsx")
        elif purpose == "tex":
            print("Usage: python xl2tex.py <input xlsx>")
            print("Example: python xl2tex.py 12.17音乐会节目信息征集.xlsx")
        else:
            raise ValueError("purpose must be either presentation or tex")
        sys.exit(1)
