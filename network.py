# %%
import pandas as pd
import numpy as np
import networkx as nx
import nxviz as nz
import matplotlib.pyplot as plt
import datetime as dt
from datetime import timedelta
import xlwings as xw

# %%
data =  pd.read_excel('network_macro.xlsm')

# %%
def find_upper_assy(x):
   
    assy=data.loc[(data['level']<x[0]) & (data.index<x.name),'material'].tail(1).to_string(index=False)
    if assy.isalpha():
        return assy
    else:
        return '-'

# %%
data['next_assy'] = data.apply(find_upper_assy,axis=1)
data

# %%
T = nx.DiGraph()  #create a directional graph
activities = data["next_assy"].unique().tolist()

T.add_nodes_from(activities)   #adding nodes

T.nodes() #view the nodes

# %%



activities_edges = [(a,b) for a,b in data[["material","next_assy"]].to_numpy()]

T.add_edges_from(activities_edges)   #add edges


T.edges()

# %%
#add weights to the existing edges
activities_weights = data["leadtime(day)"].to_list()
i=0
for s,d in T.edges():
    T[s][d]['weight'] =activities_weights[i]
    i+=1

T.edges(data=True)

# %%
#Graph the network
pos= nx.spring_layout(T)
nx.draw_networkx_nodes(T,pos)
nx.draw_networkx_edges(T,pos,arrows=True)
nx.draw_networkx_edge_labels(T,pos,edge_labels=nx.get_edge_attributes(T,'weight'))
nx.draw_networkx_labels(T,pos)
plt.savefig('network.png',dpi=100)

# %%
print(nx.shortest_path(T,'X','-',weight='weight'))
nx.shortest_path_length(T,'X','-',weight='weight')

# %%
print(nx.dag_longest_path(T,weight='weight'))
nx.dag_longest_path_length(T,weight='weight')

# %%
def find_short(x):
    return nx.shortest_path_length(T,x[1],'-',weight='weight')

# %%
data["total_time"] = nx.dag_longest_path_length(T,weight='weight')
data["impacted_time"] = data.apply(find_short,axis=1)
data

# %%
today = dt.datetime.today().strftime("%Y-%m-%d")
critic_index = data[data["impacted_time"]==data['impacted_time'].max()].index.tolist()[0]
critic_list = nx.shortest_path(T,data.iloc[critic_index][1],'-',weight='weight')
data["critical_path"]=''
data.loc[(data.index <= critic_index) & data['material'].isin(critic_list),'critical_path']='yes'

earliest_begin= max(data['start_date'].min().strftime("%Y-%m-%d"),today)
earliest_begin = dt.datetime.strptime(earliest_begin,"%Y-%m-%d")
calc_finish=  earliest_begin + timedelta(nx.dag_longest_path_length(T,weight='weight'))
calc_finish

# %%
def calc_start_date(x):
    return calc_finish- timedelta(x[8])

def calc_finish_date(x):
    return x[10] + timedelta(x[5])

# %%
data['calc_start_date'] = data.apply(calc_start_date,axis=1)
data['calc_finish_date'] = data.apply(calc_finish_date,axis=1)

# %%
wb = xw.Book("network_macro.xlsm")
ws = wb.sheets["Sheet1"]

ws['A1'].options(pd.DataFrame,header=1,index=False,expand='table').value = data

ws.autofit(axis="columns")

yes_list = data[data['critical_path']=='yes'].index.to_list()  #filter the critical path value 'yes' and find its indexes

rownum= xw.Range('B1').current_region.last_cell.row  #find the lastest row of the B column

xw.Range(f'B1:B{rownum}').color = None #remove the color of B column

for cell in yes_list:    #it changes the color to red in B column cells that are in the same index with critical path yes'es
    xw.Range(f"B{int(cell)+2}").color = '#D10000'

ws.tables.add(ws.used_range, name="a_table",table_style_name = "TableStyleLight1")  #add autofilter property to all columns
ws.tables["a_table"].show_autofilter = True

wb.save('network_macro.xlsm')

# %%
import plotly.express as px

colors= ['#D10000', '#45B08C']
fig = px.timeline(data, x_start="calc_start_date", x_end="calc_finish_date", y="material",color='critical_path',color_discrete_sequence=colors)
fig.update_yaxes(autorange="reversed") # makes the tasks are listed from up to down
fig.update_layout()
fig.write_html("gant.html")
fig.show()


# %%



