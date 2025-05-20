from openpyxl import Workbook
import pandas as pd
import ast

# 读取“线路简单”并处理 path
df = pd.read_excel("current_input.xlsx", sheet_name="线路简单")
df["path"] = df["path"].apply(ast.literal_eval)
routes = df.to_dict("records")

# 读取“揽收仓”信息并建立映射
pickup_df = pd.read_excel("current_input.xlsx", sheet_name="揽收仓")
pickup_info = pickup_df.set_index("揽收仓").to_dict("index")

# 输出表头
headers = ["揽收仓", "揽收时长（小时）", "开始时间", "开始日期", "完成时间", "完成日期", "元/公斤", "元/票"]

# 构建输出行
output_rows = []
for route in routes:
    start_point = route["path"][0]
    pickup = pickup_info.get(start_point)
    if pickup:
        row = [
            start_point,
            pickup.get("揽收时长（小时）"),
            pickup.get("开始时间"),
            pickup.get("开始日期"),
            pickup.get("完成时间"),
            pickup.get("完成日期"),
            pickup.get("元/公斤"),
            pickup.get("元/票")
        ]
        output_rows.append(row)
    else:
        print(f"⚠️ 未找到揽收仓信息：{start_point}")

# 写入到 Excel
wb = Workbook()
ws = wb.active
ws.append(headers)
for row in output_rows:
    ws.append(row)

wb.save("ut.xlsx")
print("✅ 文件已保存为 ut.xlsx")
