import os
import traceback
import message
import excel
import config

from pandas import read_csv
from pandas import ExcelWriter

from numpy import where

path = config.path

shipping_df = read_csv(os.path.join(path, 'shipping_info.csv'))
shipping_df = shipping_df[shipping_df['SalesOrigin'].isin(['PICKLIST', 'PICK', 'LOAD', 'SENT'])]
shipping_df.MasterOrder = shipping_df.MasterOrder.fillna('Others')

picking_df = read_csv(os.path.join(path, 'picking_info.csv'))

MasterOrders = list(shipping_df.MasterOrder.unique())
       
for person in MasterOrders:
    try:
        filename = os.path.join(path, '{}'.format(person), 'CDP.xlsx')
        writer = ExcelWriter(filename, engine='xlsxwriter')

        df = shipping_df[shipping_df['MasterOrder'] == person]
        sheetname = 'shipping'
        excel.make_table(df, writer, filename, sheetname)
        
        orders = list(df.SalesID)
        df = picking_df[picking_df.SalesOrderNo.apply(lambda x: True if x in orders else False)]        
        sheetname = 'picking'
        excel.make_table(df, writer, filename, sheetname)
        writer.save()
        
    except:
        var = "Person {} \n".format(person)
        var = var + traceback.format_exc()
        message.error_message(config.server, config.sender, config.receiver, var)
        raise
