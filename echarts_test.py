from pyecharts.charts import Bar
from pyecharts import options as opts
import pandas as pd 

user_table = pd.read_excel('./项目储备情况表.xlsx')
# print(user_table)
clm = user_table.columns.tolist()
# print(clm)
x_data = user_table.loc[:, clm[0]].tolist()
y_data = user_table.loc[:, clm[1]].tolist()
print(x_data)
print(y_data)

bar = Bar()
bar.add_xaxis(x_data)
bar.add_yaxis(clm[0], y_data)
bar.set_global_opts(title_opts=opts.TitleOpts(title=clm[1], subtitle="单位/万元"))
bar.render('./cbqk.html')
