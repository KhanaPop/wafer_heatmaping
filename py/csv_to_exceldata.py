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


folder_path = 'E:\TEG_TEST\241-0333(1)'
csv_files = glob(os.path.join(folder_path, '*.csv'))

combined_data = []

for file in csv_files:
    filename = os.path.basename(file)
    x, y = extract_xy_from_filename(filename)

    if x is None or y is None:
        continue  

    try:
        """df = pd.read_csv(file, header=None, delimiter=',')
        #df = pd.read_csv(file, skiprows=254, header=None ,delimiter=',') 
        print(f"{filename}: shape={df.shape}") 
        raw_value = df.iloc[255, 2]
        #print(f"{filename}: raw_value = {raw_value}")
        value = float(str(raw_value).strip().replace("'", "")) 
        #print(f"{filename}: cleaned value = {value}") 
    except Exception as e:
        value = None
        df = pd.read_csv(file, header=None, delimiter=',', engine='python') 
        print(f"{filename}: shape = {df.shape}")
        print(df.head(260).tail(5)) 
        value = df.iloc[255, 2] """
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


df_combined = pd.DataFrame(combined_data).sort_values(by=['y', 'x'])

folder_name = os.path.basename(folder_path)
output_excel = os.path.join(r'E:\test1', f"{folder_name}.xlsx")
#df_combined.to_excel(output_excel, index=False)
print(f'success: {output_excel}')
#print("data:", combined_data[:3]) 
