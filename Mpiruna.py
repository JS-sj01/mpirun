import pandas as pd
import os

files = os.listdir("MPIF")
col = row = [f'cn{str(i).rjust(3,"0")}' for i in range(59,109)]

df = pd.DataFrame(index=row,columns=col)

for f in files:
    d_col = f.split('MPI-f-')[-1]
    d_col = d_col.split('.log')[0]
    with open(os.path.join('MPIF',f),'r') as _f:
        for line in _f.readlines():
            r,data = line.split(' ')
            data = data.strip()
            d_row = r.split('ib0')[0]

            df.at[d_row,d_col] = float(data) if data != '' else 0

# df.style.applymap(lambda x:'color:red' if x%2 else "color:yellow").to_excel('analysis.xlsx','cn59-cn108')
def color_red_or_green(val):

    color = 'red' if val < 16000 else 'black'

    return 'color: %s' % color

df.style.applymap(color_red_or_green).to_excel('analysis.xlsx','cn59-cn108')
