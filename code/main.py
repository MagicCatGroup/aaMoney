import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


def open_file():
    file_path = filedialog.askopenfilename(
        title="Select an Excel file", filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        try:
            return file_path
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read the Excel file: {e}")
    else:
        messagebox.showwarning(
            "Warning", "No file selected or the selected file is not an Excel file."
        )


def get_key(dic, value):
    return [k for k, v in dic.items() if v == value]


def save_bill(bill: list[list[int]], members: list[str]) -> list[str]:
    temp = []
    for i in range(len(members)):
        for j in range(len(members)):
            if i != j and bill[i][j] > 0:
                temp.append(f"{members[i]}应付给{members[j]}:{round(bill[i][j], 2)}元")
    return temp


if __name__ == "__main__":
    # 选择excel文件
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = open_file()
    # 获取账单表单和总成员表单
    df_bill = pd.read_excel(file_path)
    df_member = pd.read_excel(file_path, sheet_name="总成员")
    # 获取成员列表与对应序号
    members = df_member.columns[1].split(" ")
    members_dict = {}
    for i in range(len(members)):
        members_dict[i] = members[i]
    # 账单矩阵 bill_matrix[i][j] 表示第 i个成员应付给第 j个成员的金额
    bill_matrix = [[0 for _ in range(len(members))] for _ in range(len(members))]
    # 开始计算账单矩阵
    bill = df_bill.iloc[1:, 2:]
    for i in range(len(bill)):
        # 计算平均金额
        amount = bill.iloc[i, 2] / len(bill.iloc[i, 1].split(" "))
        # 根据付款人与参与付款人 处理账单矩阵
        # 先获取付款人与参与付款人
        payer = bill.iloc[i, 0]
        need_payers = bill.iloc[i, 1].split(" ")
        need_payers.remove(payer)
        # 获取付款人与参与付款人的序号
        payer_index = get_key(members_dict, payer)[0]
        for j in range(len(need_payers)):
            need_payer_index = get_key(members_dict, need_payers[j])[0]
            bill_matrix[need_payer_index][payer_index] += amount
            bill_matrix[payer_index][need_payer_index] -= amount
    # 将计算的账单结果存到excel文件中 名为账单的新表中
    bill_df = pd.DataFrame(save_bill(bill_matrix, members))
    writer = pd.ExcelWriter(
        file_path, mode="a", engine="openpyxl", if_sheet_exists="new"
    )
    bill_df.to_excel(writer, sheet_name="账单", index=False, header=None)
    writer.close()
