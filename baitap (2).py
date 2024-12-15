import tkinter as tk
from tkinter import messagebox, filedialog
import csv
import pandas as pd
from datetime import datetime

# Danh sách lưu dữ liệu nhân viên
danh_sach_nhan_vien = []

# Chức năng thêm nhân viên
def them_nhan_vien():
    ma = entry_ma.get()
    ten = entry_ten.get()
    don_vi = entry_don_vi.get()
    chuc_danh = entry_chuc_danh.get()
    ngay_sinh = entry_ngay_sinh.get()
    gioi_tinh = var_gioi_tinh.get()
    
    if not ma or not ten or not ngay_sinh:
        messagebox.showerror("Lỗi", "Mã, Tên, và Ngày sinh không được để trống")
        return
    
    try:
        # Chuyển đổi ngày sinh sang định dạng datetime
        ngay_sinh_dt = datetime.strptime(ngay_sinh, "%d/%m/%Y")
    except ValueError:
        messagebox.showerror("Lỗi", "Ngày sinh không đúng định dạng DD/MM/YYYY")
        return
    
    # Thêm dữ liệu vào danh sách
    danh_sach_nhan_vien.append({
        "Mã": ma,
        "Tên": ten,
        "Đơn vị": don_vi,
        "Chức danh": chuc_danh,
        "Ngày sinh": ngay_sinh_dt,
        "Giới tính": gioi_tinh
    })
    messagebox.showinfo("Thành công", "Nhân viên đã được thêm")
    xoa_form()

# Chức năng xóa form
def xoa_form():
    entry_ma.delete(0, tk.END)
    entry_ten.delete(0, tk.END)
    entry_don_vi.delete(0, tk.END)
    entry_chuc_danh.delete(0, tk.END)
    entry_ngay_sinh.delete(0, tk.END)

# Chức năng lưu dữ liệu vào file CSV
def luu_file_csv():
    with open("nhan_vien.csv", "w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Giới tính"])
        writer.writeheader()
        for nv in danh_sach_nhan_vien:
            writer.writerow({
                "Mã": nv["Mã"],
                "Tên": nv["Tên"],
                "Đơn vị": nv["Đơn vị"],
                "Chức danh": nv["Chức danh"],
                "Ngày sinh": nv["Ngày sinh"].strftime("%d/%m/%Y"),
                "Giới tính": nv["Giới tính"]
            })
    messagebox.showinfo("Thành công", "Dữ liệu đã được lưu vào CSV")

# Chức năng lọc sinh nhật hôm nay
def sinh_nhat_hom_nay():
    hom_nay = datetime.now()
    ds_sinh_nhat = [nv for nv in danh_sach_nhan_vien if nv["Ngày sinh"].day == hom_nay.day and nv["Ngày sinh"].month == hom_nay.month]
    
    if ds_sinh_nhat:
        thong_tin = "\n".join([f"{nv['Mã']} - {nv['Tên']}" for nv in ds_sinh_nhat])
        messagebox.showinfo("Sinh nhật hôm nay", thong_tin)
    else:
        messagebox.showinfo("Sinh nhật hôm nay", "Không có nhân viên nào sinh nhật hôm nay")

# Chức năng xuất danh sách Excel
def xuat_excel():
    if not danh_sach_nhan_vien:
        messagebox.showerror("Lỗi", "Danh sách rỗng")
        return
    
    df = pd.DataFrame(danh_sach_nhan_vien)
    df["Tuổi"] = df["Ngày sinh"].apply(lambda x: (datetime.now() - x).days // 365)
    df = df.sort_values(by="Tuổi", ascending=False)
    df.to_excel("danh_sach_nhan_vien.xlsx", index=False)
    messagebox.showinfo("Thành công", "Danh sách đã được xuất ra file Excel")

# Giao diện Tkinter
root = tk.Tk()
root.title("Quản lý nhân viên")

# Widgets
frame = tk.Frame(root)
frame.pack(pady=10)

# Form nhập liệu
tk.Label(frame, text="Mã").grid(row=0, column=0)
entry_ma = tk.Entry(frame)
entry_ma.grid(row=0, column=1)

tk.Label(frame, text="Tên").grid(row=1, column=0)
entry_ten = tk.Entry(frame)
entry_ten.grid(row=1, column=1)

tk.Label(frame, text="Đơn vị").grid(row=2, column=0)
entry_don_vi = tk.Entry(frame)
entry_don_vi.grid(row=2, column=1)

tk.Label(frame, text="Chức danh").grid(row=3, column=0)
entry_chuc_danh = tk.Entry(frame)
entry_chuc_danh.grid(row=3, column=1)

tk.Label(frame, text="Ngày sinh (DD/MM/YYYY)").grid(row=4, column=0)
entry_ngay_sinh = tk.Entry(frame)
entry_ngay_sinh.grid(row=4, column=1)

var_gioi_tinh = tk.StringVar(value="Nam")
tk.Label(frame, text="Giới tính").grid(row=5, column=0)
tk.Radiobutton(frame, text="Nam", variable=var_gioi_tinh, value="Nam").grid(row=5, column=1)
tk.Radiobutton(frame, text="Nữ", variable=var_gioi_tinh, value="Nữ").grid(row=5, column=2)

# Các nút chức năng
tk.Button(frame, text="Thêm nhân viên", command=them_nhan_vien).grid(row=6, column=0)
tk.Button(frame, text="Lưu CSV", command=luu_file_csv).grid(row=6, column=1)
tk.Button(frame, text="Sinh nhật hôm nay", command=sinh_nhat_hom_nay).grid(row=7, column=0)
tk.Button(frame, text="Xuất Excel", command=xuat_excel).grid(row=7, column=1)

root.mainloop()
