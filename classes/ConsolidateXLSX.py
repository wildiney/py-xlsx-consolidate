import pandas as pd
import os
import glob
from os import path

class ConsolidateXLSX:

    def __init__(self, consolidate_from_path, save_to_path):
        self.consolidate_from_path = consolidate_from_path
        self.save_to_path = save_to_path
        self.data=''
        self.get_data()

    def get_data(self):
        print("Getting Data")
        all_data = pd.DataFrame()
        files = glob.glob(self.consolidate_from_path)
        for f in files:
            print("Loading file: ",f)
            df = pd.read_excel(f)
            df.dropna(how="all",inplace=True)
            all_data = all_data.append(df, ignore_index=True)
        self.data = all_data
        return all_data

    def save_to(self):
        if path.exists(self.save_to_path):
            print("Removing old file")
            os.remove(self.save_to_path)
        writer = pd.ExcelWriter(self.save_to_path)
        self.data.to_excel(writer,'consolidado', header=False, index=False)
        print("Creating consolidated file")
        writer.save()
        print("Finished")


