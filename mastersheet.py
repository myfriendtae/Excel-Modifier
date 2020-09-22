import os
import message
import excel
import config

from pandas import read_csv
from pandas import ExcelWriter
from pandas import merge

from excel import Copy_excel
from numpy import where

path = config.path

shipping_df = read_csv(os.path.join(path, 'shipping_info.csv'))
shipping_df = shipping_df[shipping_df['SalesOrigin'].isin(['PICKLIST', 'PICK', 'LOAD', 'SENT'])]
shipping_df.MasterOrder = shipping_df.MasterOrder.fillna('Others')

picking_df = read_csv(os.path.join(path, 'picking_info.csv'))

MasterOrders = list(shipping_df.MasterOrder.unique())

name = 'Leila'

df = shipping_df[shipping_df['MasterOrder']==name]

def main(df):
    filepath = os.path.join(path, 'Master  Sheet Template.xlsx')

    order_lists = list(df.SalesID.unique())

    for order in order_lists:
        destpath = os.path.join(path, order + '.xlsx')
        shipping_df1 = df[df.SalesID == order]

        invoice_date = list(shipping_df1.DeptartureDate)[0]
        customer_ref = list(shipping_df1.CustomerRequisition)[0]
        booking_ref = list(shipping_df1.BookingReference)[0]
        consignee = list(shipping_df1.ShipToCustomer)[0]
        trade_term = list(shipping_df1.DeliveryTerms)[0]
        dest = list(shipping_df1.ShipToFinalDestination)[0]
        payment_term = list(shipping_df1.Payment)[0]
        ship = list(shipping_df1.ShipToExportVessel)[0]
        voyage_num = list(shipping_df1.ShipToVoyageNumber)[0]
        shipping_comany = list(shipping_df1.ShipToShippinAgent)[0]
        eta_date = list(shipping_df1.ShipToETADate.astype(str))[0]
        
        spec_nums = list(shipping_df1.ItemId.unique())

        picking_df1 = picking_df[picking_df.SalesOrderNo == order]
        picking_df1.is_copy = None
        picking_df1.loc[:, 'InventoryQty'] = picking_df1['InventoryQty'].str.replace(',', '').astype(float)

        picking_df2 = picking_df1.groupby(['ContainerNo', 'SealNo'])['InventoryQty'].agg('sum').reset_index()
        picking_df2.columns = ['ContainerNo', 'SealNo', 'InventorySum']
        picking_df2.drop_duplicates(['ContainerNo', 'SealNo', 'InventorySum'], inplace=True)
        
        containers = list(picking_df2.ContainerNo)
        seals= list(picking_df2.SealNo)
        ctns = list(picking_df2.InventorySum)
        cyphers = list(picking_df1.Cypher.unique())
        prod_dates = list(picking_df1.BatchManufacturingDate.unique())
        exp_dates = list(picking_df1.ExpirationDate.unique())

        file_copy = Copy_excel(filepath, destpath)
        file_copy.write_workbook(3, 9, name)
        file_copy.write_workbook(5, 9, order)
        file_copy.write_workbook(9, 9, invoice_date)
        file_copy.write_workbook(11, 9, customer_ref)
        file_copy.write_workbook(13, 9, booking_ref)
        file_copy.write_workbook(13, 2, consignee)
        file_copy.write_workbook(26, 3, trade_term)
        file_copy.write_workbook(26, 4, dest)
        file_copy.write_workbook(27, 4, payment_term)
        file_copy.write_workbook(31, 2, ship)
        file_copy.write_workbook(31, 4, voyage_num)
        file_copy.write_workbook(31, 5, shipping_comany)
        file_copy.write_workbook(37, 6, eta_date)

        for i in range(len(spec_nums)):
            file_copy.write_workbook(39+i, 9, spec_nums[i])

        for i in range(len(containers)):
            file_copy.write_workbook(45+i, 2, containers[i])

        for i in range(len(seals)):
            file_copy.write_workbook(45+i, 5, seals[i])

        for i in range(len(ctns)):
            file_copy.write_workbook(45+i, 8, ctns[i])

        for i in range(len(cyphers)):
            file_copy.write_workbook(45+i, 12, cyphers[i])

        for i in range(len(prod_dates)):
            file_copy.write_workbook(45+i, 14, prod_dates[i])

        for i in range(len(exp_dates)):
            file_copy.write_workbook(45+i, 16, exp_dates[i])

        file_copy.save_excel()

if __name__ == '__main__':
    main(df)
            