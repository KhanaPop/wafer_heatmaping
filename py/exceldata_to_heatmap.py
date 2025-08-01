import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
#import os

# 1. ระบุ path ไปยังไฟล์ Excel ที่มีข้อมูล x, y, value
excel_path = r'E:\data\CMI_TEST_241-0333.xlsx'

# 2. โหลดข้อมูลจาก Excel
df = pd.read_excel(excel_path)

# 3. วิเคราะห์ค่าสูงสุด ต่ำสุด
min_val = df['value'].min()
max_val = df['value'].max()
print(f"Min value: {min_val}")
print(f"Max value: {max_val}")
df['abs_x'] = df['x'].abs() #convert x y value to abs
df['abs_y'] = df['y'].abs()

# 4. เตรียมข้อมูลสำหรับ Heatmap: แปลง DataFrame เป็นตารางแบบ 2 มิติ (pivot table)
pivot_table = df.pivot(index='abs_y', columns='abs_x', values='value')

# 5. วาด Heatmap
plt.figure(figsize=(10, 8))
vmin = df['value'].quantile(0.05)  # 5th percentile
vmax = df['value'].quantile(0.95)  # 95th percentile
print(f"VMin: {vmin}, VMax: {vmax}")
sns.heatmap(pivot_table, cmap='coolwarm', annot=True,annot_kws={"size": 6}, fmt=".2f", linewidths=0.5, vmin=vmin, vmax=vmax )

plt.title("Heatmap of Value at (x, y) Coordinates")
plt.xlabel("x")
plt.ylabel("y")
plt.tight_layout()
plt.show()
#plt.savefig(r'E:\test1\HEATER_heatmap.png')
