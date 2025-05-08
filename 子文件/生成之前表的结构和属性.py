import pandas as pd

import pandas as pd
import ast  # 用于将字符串安全地转换为列表

# 读取 .xlsx 文件中的特定 sheet

sheet_name = "线路简单"  # 将这里的 "Sheet1" 替换为你要读取的工作表名称

df = pd.read_excel("输入.xlsx", sheet_name=sheet_name)

# 将 path 字段的字符串转换回列表

df["path"] = df["path"].apply(ast.literal_eval)

# 将 DataFrame 转换回字典列表

routes = df.to_dict('records')

# 定义表头
headers = [
    "配送模式", "路由&资源组合", "单票1公斤", "序号", "路由", "全程距离（公里）",
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

# 创建一个空的 DataFrame，只包含表头
data = []
for i, route in enumerate(routes, start=1):
    row = {
        "配送模式": route["delivery_method"],
        "路由&资源组合": route["name"],
        "序号": i,
        "路由": " -> ".join(route["path"])
    }
    data.append(row)

df = pd.DataFrame(data, columns=headers)

# 将 DataFrame 保存为 Excel 文件
file_path = r"D:\VsCodeProjects\中乌路由pro\delivery_info.xlsx"
df.to_excel(file_path, index=False)

print(f"Excel 文件已生成，路径为: {file_path}")