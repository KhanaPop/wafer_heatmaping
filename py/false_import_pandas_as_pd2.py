import pandas as pd
import os
import re
from glob import glob

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


folder_path = 'E:\TTVmodular2x2\CMI_TEST_241-033'
csv_files = glob(os.path.join(folder_path, '*.csv'))

combined_data = []

for file in csv_files:
    filename = os.path.basename(file)
    x, y = extract_xy_from_filename(filename)

    if x is None or y is None:
        continue  

    try:
        df = pd.read_csv(file, header=None) 
        print(f"{filename}: shape={df.shape}") 
        value = float(df.iloc[255, 2])  
    except Exception as e:
        value = None

    combined_data.append({
        'filename': filename,
        'x': x,
        'y': y,
        'value': value
    })


df_combined = pd.DataFrame(combined_data).sort_values(by=['y', 'x'])


output_excel = r'E:\test0\TTVmodular2x2_CMI_TEST_241-033.xlsx'
df_combined.to_excel(output_excel, index=False)
print(f'success: {output_excel}')
