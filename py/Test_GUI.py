import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import ttk

# โหลดข้อมูล
df = pd.read_excel(r'E:\test1\D09.xlsx')

# แปลงเป็น absolute
df['abs_x'] = df['x'].abs()
df['abs_y'] = df['y'].abs()
pivot_table = df.pivot(index='abs_y', columns='abs_x', values='value')

vmin = df['value'].min()
vmax = df['value'].max()

# ฟังก์ชันสร้าง heatmap
def plot_heatmap(show_numbers):
    plt.figure(figsize=(10, 8))
    sns.heatmap(
        pivot_table, cmap='viridis', annot=show_numbers, fmt=".1f",
        linewidths=0.5, vmin=vmin, vmax=vmax,
        annot_kws={"size": 8} if show_numbers else None
    )
    plt.title("Heatmap (abs x, y)")
    plt.xlabel("abs(x)")
    plt.ylabel("abs(y)")
    plt.tight_layout()
    plt.show()

# สร้าง GUI
def start_gui():
    root = tk.Tk()
    root.title("Heatmap Viewer")

    show_value_var = tk.BooleanVar()
    show_value_var.set(True)

    # Checkbox
    checkbox = ttk.Checkbutton(root, text="แสดงค่าบน heatmap", variable=show_value_var)
    checkbox.pack(pady=10)

    # ปุ่ม Plot
    def on_plot():
        plot_heatmap(show_value_var.get())

    plot_button = ttk.Button(root, text="Plot Heatmap", command=on_plot)
    plot_button.pack(pady=10)

    root.mainloop()

# เรียก GUI
start_gui()
