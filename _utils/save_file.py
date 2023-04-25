import os
from datetime import datetime

import numpy as np
import pandas as pd
from pyexcel_ods import save_data

class Saved_file:
    def saved_file(self, filename, j, data, rep, columns):
        sheet_name="Feuille1"
        if not os.path.exists(rep):
            os.makedirs(rep)
            print("data", data[j])
        data[j][15] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        data[j][16] = 'M'
        if os.path.exists(filename):
            f = open(filename, 'a')
            f.close()
            old_data_df = pd.read_excel(filename)
            old_data = old_data_df.values.tolist()
            if old_data:
                numero_affaire = data[j][5]
                old_data = list(filter(lambda x: x[5] != numero_affaire, old_data))
                print("old data",old_data)
                print("data", data[j])
                numpyData = np.row_stack((old_data,data[j]))
                # data = list(numpyData)
                data = numpyData.tolist()
                data.insert(0, columns)
            os.remove(filename)
        else:
            data.insert(0, columns)
        save_data(filename, data=data)
