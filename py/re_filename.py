import os
import re

folder_path = r"E:\0402แยกไฟล์\R-V-RTD"

# เริ่มพิกัด
x = 0
y = 18


# ช่วงลำดับที่ต้องการเปลี่ยนชื่อ
min_index = 619
max_index = 620

# regex หาลำดับในวงเล็บ 
pattern = re.compile(r'\((\d+)\)')

# สร้างลิสต์ไฟล์ที่อยู่ในช่วงที่ต้องการ
file_list = []
for filename in os.listdir(folder_path):
    if filename.endswith(".csv"):
        match = pattern.search(filename)
        if match:
            index = int(match.group(1))
            if min_index <= index <= max_index:
                file_list.append((index, filename))

# เรียงลำดับ index จากน้อยไปมาก
file_list.sort(key=lambda x: x[0])

# เปลี่ยนชื่อไฟล์
for index, filename in file_list:
    new_filename = pattern.sub(f"_{x} {y} ({index})", filename)

    old_path = os.path.join(folder_path, filename)
    new_path = os.path.join(folder_path, new_filename)

    os.rename(old_path, new_path)
    print(f"Renamed: {filename} → {new_filename}")

    x += 1   
