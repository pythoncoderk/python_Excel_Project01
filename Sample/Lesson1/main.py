import os
import datetime

file_path = "memo/面談記録_阿部恵利様.txt"

print(file_path)

created_time = os.path.getmtime(file_path)
print(created_time)

created_time = datetime.datetime.fromtimestamp(created_time)

print(created_time)

year = created_time.year
print(year)

month = created_time.month
print(month)

day = created_time.day
print(day)

file_name = file_path.split("/")[1]
print(file_name)

new_file_name = f"{year}年{month}月{day}日_{file_name}"
print(new_file_name)

os.rename(file_path, f"memo/{new_file_name}")

print(os.path.exists("interview"))

print(os.path.exists("2020年"))

print(os.path.exists("interview/2020年"))

print(os.path.exists("interview/2021年"))

if not os.path.exists("interview/2021年/4月"):
    os.makedirs("interview/2021年/4月")

print(os.path.exists("test"))
print(not os.path.exists("test"))

if not os.path.exists("test"):
    print("testフォルダは存在しません！")

new_dir = f"interview/{year}年/{month}月"
if not os.path.exists(new_dir):
    os.makedirs(new_dir)