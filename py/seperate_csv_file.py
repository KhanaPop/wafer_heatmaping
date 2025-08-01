import os
import shutil

source_folder =r"F:\HEATER-SHORT_H8"
G1_folder =r"E:\PI25TTC9B-H8\PI25TTC9B-H8-HEATER"
G2_folder =r"E:\PI25TTC9B-H8\PI25TTC9B-H8-SHORT"

for filename in os.listdir(source_folder):
    source_path = os.path.join(source_folder, filename)
    if filename.startswith("3-wire "):
        shutil.copy(source_path, G1_folder)
    elif filename.startswith("R-V-RTD-"):
        shutil.copy(source_path, G2_folder)
print('success')        
        