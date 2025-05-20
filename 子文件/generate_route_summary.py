import pandas as pd
import ast
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
input_file = os.path.join(BASE_DIR, "..", "current_input.xlsx")
output_file = os.path.join(BASE_DIR, "..", "test_outputs", "route_summary.xlsx")

sheet_name = "线路简单"
df = pd.read_excel(input_file, sheet_name=sheet_name)
df["path"] = df["path"].apply(ast.literal_eval)
routes = df.to_dict('records')

headers = [
    "备注1", "备注2", "单票1公斤", "序号", "路由", "全程距离（公里）",
    "平台时效要求", "全程时限（天）", "全程时限（小时）", "总成本：元/公斤量纲部分",
    "总成本：元/票量纲部分", "总成本：折算至元/票", "分段成本（折算至元/票）国内段",
    "分段成本（折算至元/票）跨境干线", "分段成本（折算至元/票）清关",
    "分段成本（折算至元/票）国际经转", "分段成本（折算至元/票）末端",
    "分段成本国内段（含报关）元/公斤部分", "分段成本国内段（含报关）元/票部分",
    "分段成本跨境干线元/公斤", "分段成本国际段元/公斤部分",
    "分段成本国际段元/票部分", "国内段时限揽收（小时）",
    "国内段时限干线运输（小时）", "国内段时限转运中心等待作业（小时）",
    "国内段时限转运中心实际作业（小时）", "国内段时限转运中心实际作业（小时）",
    "国内段时限合计（天）", "跨境干线时限小时", "跨境干线时限天",
    "国际段时限干线运输（小时）", "国际段时限转运中心等待作业（小时）",
    "国际段时限转运中心实际作业（小时）", "国际段时限转运中心集货等待（小时）",
    "国际段时限末端派送（小时）", "国际段时限合计（天）"
]

data = []
for i, route in enumerate(routes, start=1):
    row = {
        "序号": i,
        "路由": " -> ".join(route["path"])
    }
    data.append(row)

df_out = pd.DataFrame(data, columns=headers)
df_out.to_excel("delivery_info.xlsx", index=False)
print("✅ 生成 delivery_info.xlsx 成功")
