import os
from datetime import datetime
import numpy as np
import pandas as pd
from pyexcel_ods import save_data
class Saved_file:
    def saved_file(self, filename, j, data, rep, columns, result):
        data_to_saved = data
        if not os.path.exists(rep):
            os.makedirs(rep)
            print("data[j]", data[j])
            data[j][14] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            data[j][15] = result

        if os.path.exists(filename):
            f = open(filename, 'a')
            f.close()
            old_data_df = pd.read_excel(filename)
            old_data = old_data_df.values.tolist()
            if old_data:
                numero_facture = data_to_saved[j][4]
                data_to_saved[j][15] = result
                old_data = list(filter(lambda x: x[4] != numero_facture, old_data))
                print("old data", len(old_data))
                print("data[j]", data_to_saved[j])
                numpyData = np.row_stack((old_data, data_to_saved[j]))
                # data = list(numpyData)
                data_to_saved = numpyData.tolist()
            os.remove(filename)
        else:
            pass

        if data_to_saved[0] == columns:
            pass
        else:
            data_to_saved.insert(0, columns)
        save_data(filename, data=data_to_saved)
