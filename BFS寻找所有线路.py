import pandas as pd
from collections import deque

# 读取 Excel 文件
file_path = r"D:\VsCodeProjects\中乌路由pro\输入.xlsx"
df = pd.read_excel(file_path, sheet_name="线路详细", usecols=["出发地", "目的地", "清关", "配送", "元/公斤", "元/票"])

# 处理 NaN，确保无空值影响路径计算
df = df.fillna("")

# **构建有向图**
graph = {}
for _, row in df.iterrows():
    start, end, clearance, delivery = row["出发地"], row["目的地"], row["清关"], row["配送"]
    price_per_kg, price_per_ticket = row["元/公斤"], row["元/票"]

    if start not in graph:
        graph[start] = set()

    # **如果有清关和配送，创建多种路径**
    if clearance and delivery:
        clearance_node = f"{end}（{clearance}清关）"
        delivery_node = f"{delivery}"
        if clearance_node not in graph:
            graph[clearance_node] = set()
        graph[start].add(clearance_node)
        graph[clearance_node].add(delivery_node)
        graph[delivery_node] = {end}
    elif clearance:
        clearance_node = f"{end}（{clearance}清关）"
        if clearance_node not in graph:
            graph[clearance_node] = set()
        graph[start].add(clearance_node)
        graph[clearance_node].add(end)
    elif delivery:
        delivery_node = f"{delivery}"
        if delivery_node not in graph:
            graph[delivery_node] = set()
        graph[start].add(delivery_node)
        graph[delivery_node].add(end)
    else:
        graph[start].add(end)

# **广度优先搜索（BFS）找到所有路径**
def find_paths_bfs(graph, start, end):
    queue = deque([[start]])  # 队列存储路径
    result_paths = set()  # **使用 `set()` 进行去重**

    while queue:
        path = queue.popleft()
        last_node = path[-1]

        if last_node == end:
            result_paths.add(tuple(path))  # **使用 `tuple()` 存储唯一路径**
            continue

        if last_node in graph:
            for next_node in graph[last_node]:
                if next_node not in path:  # **防止回路**
                    new_path = path + [next_node]
                    queue.append(new_path)

    return sorted(result_paths)  # **返回排序后的唯一路径**

# **输入起点和终点**
start_location = "霍尔果斯"
end_location = "塔什干"

# **计算所有路径**
all_routes = find_paths_bfs(graph, start_location, end_location)


# **存储结果到 lines 变量**
lines = []

# **输出结果**
if all_routes:
    print(f"从 {start_location} 到 {end_location} 的所有可能线路组合：")
    for i, route in enumerate(all_routes, start=1):
        route_str = f"{i}. {' → '.join(route)}"
        print(route_str)
        lines.append(route_str)  # 存入列表
else:
    lines.append(f"没有找到从 {start_location} 到 {end_location} 的有效线路。")
    print(lines[0])

# **此时 `lines` 变量中存储了所有线路**

##——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————

import pandas as pd

# **处理路径数据**
lines_data = []

for line in lines:
    parts = line.split(". ")  # 拆分编号和路径
    if len(parts) == 2:
        route_number = int(parts[0])  # 提取编号
        route_nodes = parts[1].split(" → ")  # 按箭头拆分路径

        # **变量初始化**
        clearance_node = ""  # 清关地点
        intermediate_node = "NA"  # 清关和目的地之间的地名，默认 "NA"
        delivery_node = ""  # 配送方式
        processed_route = []  # 普通站点
        destination = ""  # 目的地

        # **遍历路径节点**
        for i, node in enumerate(route_nodes):
            if "清关" in node:  
                clearance_node = node  # 识别清关节点
            elif "宅配" in node or "店配" in node:
                delivery_node = node  # 识别配送方式
            elif clearance_node and not delivery_node:
                # **清关地点之后的中间地名**
                intermediate_node = node
            else:
                processed_route.append(node)  # 普通站点
        
        # **设置目的地（最后一个节点）**
        if processed_route:
            destination = processed_route.pop()  # 取最后一个站点作为目的地

        # **添加到结果列表**
        lines_data.append([route_number] + processed_route + [clearance_node, intermediate_node, destination, delivery_node])

# **找出最长路径，确保 DataFrame 列数一致**
max_columns = max(len(route) for route in lines_data)

# **创建列名**
column_names = ["序号"] + [f"转运中心{i}" for i in range(max_columns - 5)] + ["清关地点", "国际转运中心", "目的地", "配送方式"]

# **转换为 DataFrame**
lines_df = pd.DataFrame(lines_data, columns=column_names)

# **填充 NaN 为 ""**
lines_df = lines_df.fillna("")

# **显示 DataFrame**
print(lines_df)

#lines_df.to_excel("lines_df.xlsx", index=False)  可选

##————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————

# **将 lines_df 转换回路径字符串列表**
lines_list = []

if not lines_df.empty:  # 如果 lines_df 里有数据
    print(f"从 {start_location} 到 {end_location} 的所有可能线路组合：")
    
    for i, row in lines_df.iterrows():
        # 过滤掉 "NA"，并连接成 " → " 形式的路径
        route_str = " → ".join(str(x) for x in row if x != "" and x != "NA")
        formatted_route = f"{i+1}. {route_str}"  # 添加编号
        print(formatted_route)  # 输出路径
        lines_list.append(formatted_route)  # 存入列表
else:
    no_route_msg = f"没有找到从 {start_location} 到 {end_location} 的有效线路。"
    lines_list.append(no_route_msg)
    print(no_route_msg)

#————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————

import pandas as pd

# 拆分“清关地点”列，提取括号前后内容
lines_df[['清关地点', '清关']] = lines_df['清关地点'].str.extract(r'(.+?)（(.+?)清关）')

# 重命名“国际转运中心”列为“转运中心4”
lines_df.rename(columns={"国际转运中心": "转运中心3"}, inplace=True)

# 删除原“清关地点”列
linesclean_df = lines_df.drop(columns=["清关地点"])

# 显示清理后的 DataFrame
print(linesclean_df)

#linesclean_df.to_excel("linesclean_df.xlsx", index=False) 可选