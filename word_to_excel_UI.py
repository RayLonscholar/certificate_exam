import tkinter as tk
from tkinter import filedialog, messagebox
from word_to_excel_process import process_word_file

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if file_path:
        file_label.config(text=f"Selected file: {file_path}", wraplength = int(app.winfo_width()/1.2))
        app.selected_file = file_path

def start_processing():
    if hasattr(app, 'selected_file'):
        process_word_file(app.selected_file)
        messagebox.showinfo("Successfully", "Processing completed and data saved to exam_data.xlsx.\n資料儲存到exam_data.xlsx")
    else:
        messagebox.showwarning("Input Required", "Please select word file.\n請選擇word檔")

app = tk.Tk()
app.title("Word to Excel User Interface")
window_width = 500
window_height = 250
app.geometry(f'{window_width}x{window_height}')

# 檔案選擇按鈕
file_btn = tk.Button(app, text="請選擇word檔案", font = ('Arial',16), command=select_file)
file_btn.pack(pady=5)
file_label = tk.Label(app, text="請選擇", font = ('Arial',10))
file_label.pack(pady=5)

# 開始處理按鈕
process_btn = tk.Button(app, text="開始轉換", font = ('Arial',16), command=start_processing)
process_btn.pack(pady=20)

app.mainloop()