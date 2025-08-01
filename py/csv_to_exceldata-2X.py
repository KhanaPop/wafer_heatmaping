import pandas as pd
import os
import re
from glob import glob


folder_path = r'Y:\XXXXX'

csv_files = glob(os.path.join(folder_path, '*.csv'))
folder = r'Z:\XXXXX'

combined_data = []

for file in csv_files:
    filename = os.path.basename(file)  

    try:
        with open(file, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()
            if len(lines) >= 256:
                line254 = lines[255].strip()
                parts = re.split(r'[,\t\s]+', line254)  # แยกด้วย comma, tab หรือ space
                if len(parts) > 2:
                    value = float(parts[2])

    except Exception as e:
        print(f"Error in {filename}: {e}")
        value = None

    combined_data.append({
        'filename': filename,
        'value': value
    })

df = pd.DataFrame(combined_data)
folder_name = os.path.basename(folder_path)
output_path = os.path.join(rf'{folder}',f"{folder_name}.xlsx")
df.to_excel(output_path, index=False)

print(f"Done: {output_path}")
