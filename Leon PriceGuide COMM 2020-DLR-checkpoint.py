# import xlwings as xw
import json
import logging

import pandas as pd
from pathlib import Path

import utils


class Leon PriceGuide COMM 2020-DLR-checkpoint:
    response = {}
    logging.basicConfig(filename='app.log', filemode='w',
                        format='%(name)s - %(levelname)s - %(message)s')

    def __init__(self, source_path, destination_path, vendor_name):
        self.source_path = source_path
        self.destination_path = destination_path
        self.vendor_name = vendor_name
        self.response["message"] = "cleared sheet successfully"
        self.response["source"] = source_path
        self.response["destination"] = destination_path

    def get_response(self):
        return self.response
        # print(self.response())

    def clean_xl(self, source_path, destination_path, response={}):
        # source_path = Path.cwd() / "data/qsc.xlsx"
        # print(source_path)
        # Create Utility class object to use common methods
        ut = utils.Utils()
        df = pd.read_csv(source_path)

        # Remove 3 rows
        # df.drop(index=[0, 1, 2])
        # df = df.iloc[3:]

        # Removing unwanted columns
        df1 = pd.DataFrame(df, index=range(0, 78))
        df1.dropna(thresh=2, inplace=True)
        df1.dropna(axis=1, how='all', inplace=True)
        df1.columns = df1.iloc[0]
        df1.drop(1, inplace=True)
        df2 = pd.DataFrame(df, index=range(78, 87))
        df2.dropna(thresh=2, inplace=True)
        df2.dropna(axis=1, how='all', inplace=True)
        df2.columns = df2.iloc[0]
        df2.drop(78, inplace=True)
        df3 = pd.DataFrame(df, index=range(87, 156))
        df3.dropna(thresh=7, inplace=True)
        df3.dropna(axis=1, how='all', inplace=True)
        df3.columns = df3.iloc[0]
        df3.drop(87, inplace=True)
        df4 = pd.DataFrame(df, index=range(156, 171))
        df4.dropna(thresh=2, inplace=True)
        df4.dropna(axis=1, how='all', inplace=True)
        df4.columns = df4.iloc[0]
        df4.drop(156, inplace=True)
        df5 = pd.DataFrame(df, index=range(171, 189))
        df5.dropna(thresh=7, inplace=True)
        df5.dropna(axis=1, how='all', inplace=True)
        df5 = df5[df5["Dealer "].str.match('[^a-zA-Z]')]
        df6 = pd.DataFrame(df, index=range(189, 194))
        df6.dropna(thresh=2, inplace=True)
        df6.dropna(axis=1, how='all', inplace=True)
        df6.columns = df6.iloc[0]
        df6.drop(189, inplace=True)
        df7 = pd.DataFrame(df, index=range(194, 209))
        df7.dropna(thresh=3, inplace=True)
        df7.dropna(axis=1, how='all', inplace=True)
        df7 = df7[df7["Width "].str.match('[^a-zA-Z]')]
        df8 = pd.DataFrame(df, index=range(209, 237))
        df8.dropna(thresh=3, inplace=True)
        df8.dropna(axis=1, how='all', inplace=True)
        df8.columns = df8.iloc[0]
        df8.drop(209, inplace=True)
        # Remove empty rows
        # df.dropna(thresh=4, inplace=True)
        # df.dropna(axis=1, how='all', inplace=True)

        # df['Dealer '] = df['Dealer '].str.replace('$', '')

        # df = df[pd.to_numeric(df['Dealer '], errors='coerce').notnull()]
        #
        # for x in range(0, 1):
        #     df['Dealer '] = "$" + df['Dealer ']

        # df = df.fillna(method='ffill')
        # df = df.dropna(thresh=2)

        # Remove first column if we are not using usecols in read_excel method of pd
        # df = df.iloc[:, 1:]

        # Export to CSV
        # df.to_csv(Path.cwd() / "data/cleaned.csv", index=False)
        df.to_csv(destination_path, header=True, sep=',', index=False)

        response['message'] = "Cleaned Xls"
        response['status'] = "success"
        response['filePath'] = destination_path
        # print(df)
        print(json.dumps(response))
