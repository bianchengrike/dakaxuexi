from pyecharts.charts import Bar,Map
from pyecharts import options as opts
import pandas as pd 

user_table = pd.read_excel('./项目储备情况表.xlsx')
# print(user_table)
clm = user_table.columns.tolist()
# print(clm)
x_data = user_table.loc[:, clm[0]].tolist()
y_data = user_table.loc[:, clm[1]].tolist()
y2_data = user_table.loc[:, clm[2]].tolist()

# 生成单柱
# bar = Bar()
# bar.add_xaxis(x_data)
# bar.add_yaxis(clm[1], y_data)
# bar.set_global_opts(title_opts=opts.TitleOpts(title=clm[1]))
# bar.render('./cbqk.html')

# 生产2个柱子
# bar = Bar()
# bar.add_xaxis(x_data)
# bar.add_yaxis(clm[1], y_data)
# bar.add_yaxis(clm[2], y2_data)
# bar.set_global_opts(title_opts=opts.TitleOpts(title=clm[1]))
# bar.render('./cbqk2.html')


map=Map()
map.add(clm[1], [list(z) for z in zip(x_data, y_data)], "china",is_map_symbol_show=False)
map.set_global_opts(title_opts=opts.TitleOpts(title="Map-基本示例"), visualmap_opts=opts.VisualMapOpts(max_=5000))
map.render("map.html")
