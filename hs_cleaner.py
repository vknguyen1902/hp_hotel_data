import pandas as pd
import numpy as np
!pip install xlrd

'''
Input:
+ inputfile (type: str): path to Excel file
+ sheetname (type: str): sheet to be cleaned by the function

Output: 
+ type: None
+ Post-condition: all the empty and extraneous rows will be dropped and 
    the resulting data will be saved to a new CSV file
'''
def hs_cleaner(inputfile, sheetname):
    xl = pd.ExcelFile(inputfile)
    df = xl.parse(sheetname)

    # Skip the first 4 lines (file title and empty lines)
    df = df[4:] 

    df.columns = ["id", "bill_number", "room_number", "room_type", "checkin", "checkin_time", \
                    "checkin_date", "checkout", "checkout_time", "checkout_date", "by_day", "by_hour", \
                    "by_night", "room_price", "service_fee", "laundry_fee", "phone_card_fee", "vat", \
                    "other_fee", "other_discount", "total_fee", "paid_before", "ppaid_before", "credit", \
                    "credit_payment", "expense", "actual_fee_payment", "note", "random", "num_night_check_in", \
                    "num_morning_check_in", "num_by_hour", "num_by_night", ""]

    trash = [] # store all the row index to be dropped
    trashline = 0 # store the last line to read
    num_rows = df.shape[0]
    # Loop over each row, store the indices to be removed
    for i in range(num_rows):
        # Everything after this row will be dropped
        if df.iloc[i]["id"] == "TỔNG HỢP NỢ, TRẢ, TRẢ TRƯỚC, TRỪ TRẢ TRƯỚC TRONG NGÀY":
            trashline = i
            break

        # If row doesn't have bill number, there is no transaction going on 
        # so it should be dropped too
        if df.iloc[i]["bill_number"] is np.nan:
            trash.append(i)        

    # Drop the trashy rows
    print("Rows to be dropped:", trash, "and", trashline)
    df = df.drop(df.index[[j for j in range(trashline, num_rows)]])
    df = df.drop(df.index[trash])

    # Save dataframe to csv
    filename = inputfile[:-5] + "_cleaned" + ".csv"
    df.to_csv(filename, sep=',', encoding='utf-8', index=False)


def main():
    inputfile = "HP 010318.xlsx"
    sheetname = "THHP"
    hs_cleaner(inputfile, sheetname)

    # TODO: walk directories

if __name__ == '__main__':
    main()
