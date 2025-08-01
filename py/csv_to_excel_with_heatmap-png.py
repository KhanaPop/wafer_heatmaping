import pandas as pd
import os
import re
from glob import glob
import matplotlib.pyplot as plt
import seaborn as sns

def extract_xy_from_filename(filename):
    """
    ดึงค่าพิกัด x y จากชื่อไฟล์ที่อยู่ในรูปแบบ ..._x y... เช่น _-35 -69
    """
    match = re.search(r'_(-?\d+)\s(-?\d+)', filename)
    if match:
        x = int(match.group(1))
        y = int(match.group(2))
        return x, y
    return None, None


folder_path = 'X:\XXX'
csv_files = glob(os.path.join(folder_path, '*.csv'))

combined_data = []

for file in csv_files:
    filename = os.path.basename(file)
    x, y = extract_xy_from_filename(filename)

    if x is None or y is None:
        continue  

    try:
        
        with open(file, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()
            if len(lines) >= 256:
                line256 = lines[255].strip()
                parts = re.split(r'[,\t\s]+', line256)  # แยกด้วย comma, tab หรือ space
                if len(parts) > 2:
                    value = float(parts[2])
    except Exception as e:
        print(f"Error in {filename}: {e}")
        value = None

    combined_data.append({
        'filename': filename,
        'x': x,
        'y': y,
        'value': value
    })


df_combined = pd.DataFrame(combined_data).sort_values(by=['x', 'y'])

min_val = df_combined['value'].min()
max_val = df_combined['value'].max()
print(f"Min value: {min_val}")
print(f"Max value: {max_val}")

df_combined['abs_x'] = df_combined['x'].abs() #convert x y value to abs
df_combined['abs_y'] = df_combined['y'].abs()

folder_name = os.path.basename(folder_path)
output_excel = os.path.join(r'Z:\XXXXX', f"{folder_name}.xlsx")
df_combined.to_excel(output_excel, index=False) #convert dataframe to excel
print(f'success: {output_excel}')
#print("data:", combined_data[:3]) 

pivot_table = df_combined.pivot(index='abs_y', columns='abs_x', values='value') #plot heat map

plt.figure(figsize=(10, 8))
vmin = df_combined['value'].quantile(0.05)  # 5th percentile
vmax = df_combined['value'].quantile(0.95)  # 95th percentile
print(f"VMin: {vmin}, VMax: {vmax}")
sns.heatmap(pivot_table, cmap='coolwarm', annot=False,annot_kws={"size": 6}, fmt=".2f", linewidths=0.5, vmin=vmin, vmax=vmax )

plt.title(f"Heatmap of {folder_name} at (x, y) Coordinates")
plt.xlabel("x")
plt.ylabel("y")
plt.tight_layout()
plt.show()