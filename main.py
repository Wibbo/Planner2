from task import DAITask
from tabulate import tabulate
import time

DAI_task_list = DAITask("DAI4.xlsx")
df_tasks = DAI_task_list.df_tasks

print(tabulate(df_tasks, headers='keys'))

df_tasks.to_excel('RawData.xlsx')


