from logging import getLogger
from logging import StreamHandler
from logging import Formatter
import pandas as pd
import argparse
from pathlib import Path
from datetime import datetime
import sys


license = """
MIT License

Copyright (c) 2024 Wataru Uegami <wuegami@gmail.com>

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""


def show_welcome_message():
    print()
    print("==================================")
    print("Welcome to the IHC billing system!")
    print("==================================")

    print(f"{license}")
    print("Last update: 2024-11-17")
    print()


def read_master_file(master_file: Path, logger):
    logger.info(f"master_file: {master_file}")

    try:
        master_df = pd.read_excel(master_file, sheet_name="master", index_col="item")
        blacklist = pd.read_excel(master_file, sheet_name="IHC_blacklist")
        other = pd.read_excel(master_file, sheet_name="other", index_col='key')
        blacklist = blacklist["name"].tolist()
        logger.info(f"IHC blacklist: {blacklist}")

    except FileNotFoundError as e:
        logger.error("Master file does not exist")
        return pd.DataFrame()

    except Exception as e:
        logger.error(e)
        return pd.DataFrame()

    logger.info("Master file was read successfully")

    return master_df, blacklist, other


def main():
    show_welcome_message()
    logger = getLogger(__name__)

    stream_handler = StreamHandler()
    formatter = Formatter("%(asctime)s|%(name)s|%(levelname)s|%(message)s")
    stream_handler.setLevel("INFO")
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    # log level
    logger.setLevel("INFO")

    parser = argparse.ArgumentParser()
    parser.add_argument("input_file")
    args = parser.parse_args()

    master_file = Path("./master/master.xlsx")
    master_df, omit_IHC_list, other = read_master_file(master_file, logger)

    print(other)
    consumption_tax = other.loc['tax_rate', 'val']

    # Institute selection -------------
    master_df_columns = master_df.columns
    master_df_columns = [
        col
        for col in master_df_columns
        if col not in ["fee", "IHC1", "IHC2", "highlight"]
    ]

    print()
    print("Choose the Institute")

    if len(master_df_columns) == 0:
        logger.error("No Institute in the master file. Please check the master file.")
        sys.exit(1)

    for idx, col in enumerate(master_df_columns):
        print(f"{idx}: {col}")

    institute_idx = input("Institute: ")

    try:
        institute_name = master_df_columns[int(institute_idx)]
        print()
        logger.info(f"{institute_name}, Selected!")

    except Exception as e:
        logger.error(f"Invalid input:{e}, {institute_idx}, Aborted...")
        sys.exit(2)

    # Raw file from LIS -------------
    input_file = Path(args.input_file)
    logger.info(f"input_file: {input_file}")

    if input_file.suffix != ".xlsx":
        logger.error(f"{input_file} is not xlsx format")
        return

    # read the xlsx file
    try:
        input_df = pd.read_excel(input_file)

    except FileNotFoundError as e:
        logger.error(f"{input_file} does not exist")
        return

    except Exception as e:
        logger.error(f"Unknown error: {e}")
        return

    item_highlight: dict = master_df.to_dict()["highlight"]
    item_IHC1_dict: dict = master_df.to_dict()["IHC1"]
    item_IHC2_dict: dict = master_df.to_dict()["IHC2"]

    output1 = {}
    output2 = {}

    special_IHCs = set(master_df["IHC1"].tolist() + master_df["IHC2"].tolist())

    # For each case,
    for case_idx, row in input_df.iterrows():
        logger.info(f'* Processing {case_idx+1}/{len(input_df)}: {row["標本番号"]}')

        case_id = row["標本番号"]
        stain_list = row["染色名"].split(",")

        # delete non-IHC from the list (inplace), if it exists
        # because is should not be counted as IHC
        for omit in omit_IHC_list:
            if omit in stain_list:
                stain_list.remove(omit)

        detail_1 = {}
        detail_2 = {}

        if "材料数" in row.index:
            detail_1['材料数'] = row['材料数']

        count_more_than_3 = 0

        for billing_item in master_df.index:

            IHC_ref_1 = item_IHC1_dict[billing_item]
            IHC_ref_2 = item_IHC2_dict[billing_item]
            highlight = item_highlight[billing_item]

            include = []
            not_include = []

            if not pd.isnull(highlight):

                if highlight == "ク":
                    IHC_list_other = list(set(stain_list) - special_IHCs)

                    if len(IHC_list_other) > 0:
                        detail_1[billing_item] = 1

                        if len(IHC_list_other) > 3:
                            more_than_3 = IHC_list_other[3:]
                            IHC_list_other = IHC_list_other[:3]

                            detail_2[
                                "注１（３）ケ以外の免疫染色標本を作製した場合、４抗体目から１抗体につき"
                            ] = ",".join(more_than_3)
                            count_more_than_3 = len(more_than_3)

                        IHC_list_other = ",".join(IHC_list_other)
                        detail_2["ク"] = IHC_list_other

            if "ケ以外の免疫染色標本を作製した場合" in billing_item:
                if count_more_than_3 > 0:
                    detail_1[billing_item] = count_more_than_3
                    continue
                else:
                    detail_1[billing_item] = 0

            for IHC in [IHC_ref_1, IHC_ref_2]:
                if type(IHC) != str:
                    continue
                if IHC.startswith("_"):
                    not_include.append(IHC[1:])
                    continue
                include.append(IHC)

            if len(include) == 0:
                detail_1[billing_item] = 0
                continue

            logger.debug(f"{case_id}: {billing_item=}, {include=}, {not_include=}")

            # if include is all in stain_list, AND, not_include is not in stain_list
            if all([IHC in stain_list for IHC in include]) and all(
                [IHC not in stain_list for IHC in not_include]
            ):
                detail_1[billing_item] = 1

                if not pd.isnull(highlight):
                    detail_2[highlight] = ",".join(include)
            else:
                detail_1[billing_item] = 0

        output1[case_id] = detail_1
        output2[case_id] = detail_2

    output_df = pd.DataFrame(output1)
    output_df["請求件数"] = output_df.sum(axis=1)

    output_df["検査料(税別)"] = master_df["fee"]
    output_df["委託割合"] = master_df[master_df_columns[int(institute_idx)]]
    output_df["検査料金"] = output_df["検査料(税別)"] * output_df["委託割合"]
    output_df["金額（税別）"] = output_df["検査料金"] * output_df["請求件数"]

    # sum up the total
    total_fee = output_df["金額（税別）"].sum()
    logger.info(f"Total fee:  {total_fee} JPY")

    tax = total_fee * consumption_tax
    logger.info(f"Consumption tax: {tax} JPY")

    total_fee_with_tax = total_fee + tax
    logger.info(f"Total fee with tax: {total_fee_with_tax} JPY")

    output_3 = pd.DataFrame({
        'ご請求金額': [total_fee_with_tax],
        '病理検査料等': [total_fee],
        "標本送付料": [0],
        '税抜金額合計': [total_fee],
        '消費税': [tax],
    }).T

    output2 = pd.DataFrame(output2)

    output = pd.concat([output_df, output2], axis=0)

    output = output.replace(0, "")

    time_now = datetime.now().strftime("%Y%m%d%H%M")
    # output_filename = f"out_{institute_name}_from_{input_file.name}_{time_now}.csv"
    # output.to_csv(output_filename, encoding="cp932")

    # to xlsx
    # sheet1: output3
    # sheet2: output
    output_filename = f"out_{institute_name}_from_{input_file.name}_{time_now}.xlsx"
    with pd.ExcelWriter(output_filename) as writer:
        output_3.to_excel(writer, sheet_name="合計", header=False)
        output.to_excel(writer, sheet_name="詳細", header=True)

    logger.info(f"{output_filename} was created successfully.")
    logger.info(f"Finished.")


if __name__ == "__main__":
    main()
