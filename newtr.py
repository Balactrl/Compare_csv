import pandas as pd
from sqlalchemy import create_engine
import openpyxl

# Read server details from an Excel file
server_details_df = pd.read_excel("server_details3.xlsx")

# Create an Excel writer object to export the data
excel_file = "output.xlsx"
excel_writer = pd.ExcelWriter(excel_file, engine='openpyxl')

# Define a list of sheet names corresponding to your queries
sheet_names = ["PRT", "FRT", "AI", "AR", "BTI", "BTR", "PRC","FRC","DWR","OP","OPR","NMRC","ADJ","RJ"]  # Add more as needed

# Create a dictionary to keep track of the last row written to each sheet
sheet_last_row = {}

# Iterate through servers
for index, row in server_details_df.iterrows():
    server_name = row['ServerName']
    connection_string = row['ConnectionString']

    try:
        # Define your database connection using the extracted connection string
        engine = create_engine(connection_string)

        # Iterate through queries
        for i, sheet_name in enumerate(sheet_names):
            df = None

             # Execute the corresponding SQL query based on the sheet name
            if sheet_name == "PRT":
                sql_query = "Select STORE,(select name from ax.INVENTSITE where store=SITEID) name,(select acxclustercode from ax.INVENTSITE  where SITEID in (select store from CONTROL where REG_NUM='001'))cluster,TOSTORE,TRANSACTIONID,TRANSDATE,sum(qty) qty from AXPOSSTORETRANSFER where tostore like '%vP%' and TRANSDATE>='2024-08-01' and TRANSDATE<'2024-09-01' group by STORE,TOSTORE,TRANSACTIONID,TRANSDATE"
            elif sheet_name == "FRT":
                sql_query = "Select STORE,(select name from ax.INVENTSITE where store=SITEID) name,(select acxclustercode from ax.INVENTSITE  where SITEID in (select store from CONTROL where REG_NUM='001'))cluster,TOSTORE,TRANSACTIONID,TRANSDATE,sum(qty) qty from AXPOSSTORETRANSFER where tostore like '%VF%'and TRANSDATE>='2024-08-01' and TRANSDATE<'2024-09-01' group by STORE,TOSTORE,TRANSACTIONID,TRANSDATE"
            elif sheet_name == "AI":
                sql_query = "SELECT AH.SOURCESTORE SITEID,(select name from ax.INVENTSITE where AH.SOURCESTORE=SITEID) name,(select acxclustercode from ax.INVENTSITE  where SITEID in (select store from CONTROL where REG_NUM='001'))cluster,CAST(POSTINGDATE AS DATE) TRXDATE,SUBSTRING(POSTEDDOCNO,1,LEN(POSTEDDOCNO)) TRNO,SUM(cast(QTY*MRP as numeric(12,2))) ITEMVALUE,REMARKS FROM AX.ACXINVENTORYADJUSTMENT AH ,AX.ACXINVENTORYADJUSTMENTLINE AL WHERE AH.INTERNALDOCNO = AL.INTERNALDOCNO AND AH.SOURCESTORE =AL.SOURCESTORE AND ISPOSTED=1  and Qty < 0  and cast(POSTINGDATE as date) BETWEEN '01-aug-2024' and '01-sep-2024' GROUP BY  CAST(POSTINGDATE AS DATE), POSTEDDOCNO, REMARKS,AH.SOURCESTORE"
            elif sheet_name == "AR":
                sql_query = "SELECT AH.SOURCESTORE SITEID,(select name from ax.INVENTSITE where AH.SOURCESTORE=SITEID) name,(select acxclustercode from ax.INVENTSITE  where SITEID in (select store from CONTROL where REG_NUM='001'))cluster,CAST(POSTINGDATE AS DATE) TRXDATE,SUBSTRING(POSTEDDOCNO,1,LEN(POSTEDDOCNO)) TRNO,SUM(cast(QTY*MRP as numeric(12,2))) ITEMVALUE,REMARKS FROM AX.ACXINVENTORYADJUSTMENT AH ,AX.ACXINVENTORYADJUSTMENTLINE AL WHERE AH.INTERNALDOCNO = AL.INTERNALDOCNO AND AH.SOURCESTORE =AL.SOURCESTORE AND ISPOSTED=1  and Qty > 0  and cast(POSTINGDATE as date) BETWEEN '01-aug-2024' and '01-sep-2024'GROUP BY  CAST(POSTINGDATE AS DATE), POSTEDDOCNO, REMARKS,AH.SOURCESTORE"
            elif sheet_name == "BTI":
                sql_query = "Select  STORE,(select name from ax.INVENTSITE where store=SITEID) name,(select acxclustercode from ax.INVENTSITE  where SITEID in (select store from CONTROL where REG_NUM='001'))cluster,cast(Invoicedate as date) TRXDATE,TOSTORE SITEID,dbo.get_storename(TOSTORE) NAME, SUBSTRING(transactionid,1,LEN(transactionid)) TRNO,isnull( sum(Cast(qty as numeric)),0)QTY, isnull(sum(Cast(qty*mrp as numeric(12,2))),0) ITEMVALUE from AXPOSStoretransfer a Left Join (select  C.ITEMID,C.INVENTBATCHID,isnull(ibe.Maximumretailprice_in,c.MRP) MRP from INVENTBATCH c LEFT JOIN ax.acxinventbatchexception ibe ON ibe.itemid = c.itemid and ibe.inventbatchid =c.inventbatchid and STATEID = 'KA') as b on a.itemid=b.itemid and a.batchno=b.inventbatchid WHERE  transactionid like '%TRT%'  AND cast(Invoicedate as date) between '01-aug-2024' AND '01-sep-2024'group by STORE,cast(Invoicedate as date),TOSTORE,transactionid order by 1"
            elif sheet_name == "BTR":
                sql_query = "select INVENTSITEID SITEID,(select name from ax.INVENTSITE where INVENTSITEID=SITEID) name,(select acxclustercode from ax.INVENTSITE  where SITEID in (select store from CONTROL where REG_NUM='001'))cluster,cast(datephysical as date) TRXDATE,case when referenceid like 'JN%' THEN 'ADJUSTMENT' WHEN FROMINVENTSITEID LIKE '149%' THEN FROMINVENTSITEID WHEN FROMINVENTSITEID='' THEN VENDGROUP  WHEN FROMINVENTSITEID NOT LIKE '149%' THEN dbo.get_storename(FROMINVENTSITEID) end NAME,referenceid TRNO,VOUCHERPHYSICAL,Cast(sum(qty) as numeric(12))QTY,(Cast(sum(qty*mrp) as numeric(12,2)))ITEMVALUE,case when referenceid like 'JN%' THEN 'ADJ' WHEN FROMINVENTSITEID LIKE '149%' THEN FROMINVENTSITEID WHEN FROMINVENTSITEID='' THEN VENDGROUP WHEN FROMINVENTSITEID NOT LIKE '149%' THEN 'BTREC' end [TYPE]  from ax.acxinventtranslocal  a where DATEPHYSICAL BETWEEN '01-aug-2024' and '01-sep-2024' and VENDGROUP !='AHL_PHARMA' AND VENDGROUP !='AHL_FMCG'and referenceid NOT like 'JN%'and a.FROMINVENTSITEID NOT LIKE '%v%' and VENDGROUP NOT IN ('OP') group by referenceid ,VOUCHERPHYSICAL, cast(DATEPHYSICAL as date),FROMINVENTSITEID,vendgroup,INVENTSITEID ORDER BY 2"
            elif sheet_name == "PRC":
                sql_query = "select INVENTSITEID SITEID,(select name from ax.INVENTSITE where INVENTSITEID=SITEID) name,(select acxclustercode from ax.INVENTSITE  where SITEID in (select store from CONTROL where REG_NUM='001'))cluster,cast(datephysical as date) TRXDATE,case when referenceid like 'JN%' THEN 'ADJUSTMENT' WHEN FROMINVENTSITEID LIKE '149%' THEN FROMINVENTSITEID WHEN FROMINVENTSITEID='' THEN VENDGROUP  WHEN FROMINVENTSITEID NOT LIKE '149%' THEN dbo.get_storename(FROMINVENTSITEID) end NAME,referenceid TRNO,Cast(sum(qty) as numeric(12))QTY,(Cast(sum(qty*mrp) as numeric(12,2)))ITEMVALUE,case when referenceid like 'JN%' THEN 'ADJ' WHEN FROMINVENTSITEID LIKE '149%' THEN FROMINVENTSITEID WHEN FROMINVENTSITEID='' THEN VENDGROUP WHEN FROMINVENTSITEID NOT LIKE '149%' THEN 'BTREC' end [TYPE]  from ax.acxinventtranslocal  a where DATEPHYSICAL BETWEEN '01-aug-2024' and '01-sep-2024' and VENDGROUP='AHL_PHARMA'group by referenceid , cast(DATEPHYSICAL as date),FROMINVENTSITEID,vendgroup,INVENTSITEID ORDER BY 2"
            elif sheet_name == "FRC":
                sql_query = "select INVENTSITEID SITEID,(select name from ax.INVENTSITE where INVENTSITEID=SITEID) name,(select acxclustercode from ax.INVENTSITE  where SITEID in (select store from CONTROL where REG_NUM='001'))cluster,cast(datephysical as date) TRXDATE,case when referenceid like 'JN%' THEN 'ADJUSTMENT' WHEN FROMINVENTSITEID LIKE '149%' THEN FROMINVENTSITEID WHEN FROMINVENTSITEID='' THEN VENDGROUP  WHEN FROMINVENTSITEID NOT LIKE '149%' THEN dbo.get_storename(FROMINVENTSITEID) end NAME,referenceid TRNO,Cast(sum(qty) as numeric(12))QTY,(Cast(sum(qty*mrp) as numeric(12,2)))ITEMVALUE,case when referenceid like 'JN%' THEN 'ADJ' WHEN FROMINVENTSITEID LIKE '149%' THEN FROMINVENTSITEID WHEN FROMINVENTSITEID='' THEN VENDGROUP WHEN FROMINVENTSITEID NOT LIKE '149%' THEN 'BTREC' end [TYPE]  from ax.acxinventtranslocal  a where DATEPHYSICAL BETWEEN '01-aug-2024' and '01-sep-2024' and VENDGROUP='AHL_FMCG'group by referenceid , cast(DATEPHYSICAL as date),FROMINVENTSITEID,vendgroup,INVENTSITEID ORDER BY 2"
            elif sheet_name == "DWR":
                sql_query = "Select SUBSTRING(transactionid,1,LEN(transactionid)) TRNO,TOSTORE DCCODE,format(INVOICEDATE,'dd-MM-yyyy') TRXDATE,isnull(Cast(sum(qty*mrp) as numeric(12,2)),0)ITEMVALUE ,VendACCOUNT as VENDORCODE,'' as REMARKS from AXPOSStoretransfer as a Left Join (select  C.ITEMID,C.INVENTBATCHID,isnull(ibe.Maximumretailprice_in,c.MRP) MRP,format(expdate,'MMyy') as EXPDATE from INVENTBATCH c LEFT JOIN ax.acxinventbatchexception ibe ON ibe.itemid = c.itemid and ibe.inventbatchid =c.inventbatchid and STATEID = 'KA') as b on a.itemid=b.itemid and a.batchno=b.inventbatchid WHERE DOCUMENTNO like 'PO%' AND cast(Invoicedate as date) Between  '01-aug-2024' and '01-sep-2024'group by transactionid,TOSTORE,INVOICEDATE,VendACCOUNT Order BY 3"
            elif sheet_name == "OP":
                sql_query = "select INVENTSITEID SITEID,(select name from ax.INVENTSITE where INVENTSITEID=SITEID) name,(select acxclustercode from ax.INVENTSITE  where SITEID in (select store from CONTROL where REG_NUM='001'))cluster,cast(datephysical as date) TRXDATE,case when referenceid like 'JN%' THEN 'ADJUSTMENT' WHEN FROMINVENTSITEID LIKE '149%' THEN FROMINVENTSITEID WHEN FROMINVENTSITEID='' THEN VENDGROUP  WHEN FROMINVENTSITEID NOT LIKE '149%' THEN dbo.get_storename(FROMINVENTSITEID) end NAME,referenceid TRNO,Cast(sum(qty) as numeric(12))QTY,(Cast(sum(qty*mrp) as numeric(12,2)))ITEMVALUE,case when referenceid like 'JN%' THEN 'ADJ' WHEN FROMINVENTSITEID LIKE '149%' THEN FROMINVENTSITEID WHEN FROMINVENTSITEID='' THEN VENDGROUP WHEN FROMINVENTSITEID NOT LIKE '149%' THEN 'BTREC' end [TYPE]  from ax.acxinventtranslocal  a where DATEPHYSICAL BETWEEN '01-aug-2024' and '01-sep-2024' and VENDGROUP ='OP'group by referenceid , cast(DATEPHYSICAL as date),FROMINVENTSITEID,vendgroup,INVENTSITEID ORDER BY 2"
            elif sheet_name == "OPR":
                sql_query = "select STORE,(select name from ax.INVENTSITE where store=SITEID) name,(select acxclustercode from ax.INVENTSITE  where SITEID in (select store from CONTROL where REG_NUM='001'))cluster,cast(invoicedate as date)TRXDATE ,DOCUMENTNO,REFERENCENO TRNO,VENDORCODE,Cast(sum(qty) as numeric(12))QTY,(Cast(sum(qty*mrp) as numeric(12,2)))ITEMVALUE from inventbatch a,AXPOSPURCHASERETURN b  where invoicedate between '01-aug-2024' and '01-sep-2024'and  a.itemid=b.itemid and a.INVENTBATCHID=b.batchno and DOCUMENTNO like '%OPR%'group by DOCUMENTNO,REFERENCENO,VENDORCODE,STORE,invoicedate  order by 1,2,6"
            elif sheet_name == "NMRC":
                sql_query = "select store,(select name from ax.INVENTSITE where store=SITEID) name,(select acxclustercode from ax.INVENTSITE  where SITEID in (select store from CONTROL where REG_NUM='001'))cluster,cast(invoicedate as date)invdate,tostore,(select name from ax.INVENTSITE where tostore=SITEID) name,DOCUMENTNO,TRANSACTIONID,AHVENDACCOUNT from AXPOSSTORETRANSFER where INVOICEDATE>='2024-08-01' and INVOICEDATE<'2024-09-01'and DOCUMENTNO not like '%rt%'group by invoicedate,store,tostore,DOCUMENTNO,TRANSACTIONID,AHVENDACCOUNT"       
            elif sheet_name == "ADJ":
                sql_query = "select INVENTSITEID SITEID,(select name from ax.INVENTSITE where INVENTSITEID=SITEID) name,(select acxclustercode from ax.INVENTSITE  where SITEID in (select store from CONTROL where REG_NUM='001'))cluster,cast(datephysical as date) TRXDATE,case when referenceid like 'JN%' THEN 'ADJUSTMENT' WHEN FROMINVENTSITEID LIKE '149%' THEN FROMINVENTSITEID WHEN FROMINVENTSITEID='' THEN VENDGROUP  WHEN FROMINVENTSITEID NOT LIKE '149%' THEN dbo.get_storename(FROMINVENTSITEID) end NAME,referenceid TRNO,Cast(sum(qty) as numeric(12))QTY,(Cast(sum(qty*mrp) as numeric(12,2)))ITEMVALUE,case when referenceid like 'JN%' THEN 'ADJ' WHEN FROMINVENTSITEID LIKE '149%' THEN FROMINVENTSITEID WHEN FROMINVENTSITEID='' THEN VENDGROUP WHEN FROMINVENTSITEID NOT LIKE '149%' THEN 'BTREC' end [TYPE]  from ax.acxinventtranslocal  a where DATEPHYSICAL BETWEEN '01-aug-2024' and '01-sep-2024' and referenceid like 'JN%'group by referenceid , cast(DATEPHYSICAL as date),FROMINVENTSITEID,vendgroup,INVENTSITEID ORDER BY 2"                              
            elif sheet_name == "RJ":
                sql_query = "select INVENTSITEID SITEID,(select name from ax.INVENTSITE where INVENTSITEID=SITEID) name,(select acxclustercode from ax.INVENTSITE  where SITEID in (select store from CONTROL where REG_NUM='001'))cluster,cast(datephysical as date) TRXDATE,case when referenceid like 'JN%' THEN 'ADJUSTMENT' WHEN FROMINVENTSITEID LIKE '149%' THEN FROMINVENTSITEID WHEN FROMINVENTSITEID='' THEN VENDGROUP  WHEN FROMINVENTSITEID NOT LIKE '149%' THEN dbo.get_storename(FROMINVENTSITEID) end NAME,referenceid TRNO,Cast(sum(qty) as numeric(12))QTY,(Cast(sum(qty*mrp) as numeric(12,2)))ITEMVALUE,case when referenceid like 'JN%' THEN 'ADJ' WHEN FROMINVENTSITEID LIKE '149%' THEN FROMINVENTSITEID WHEN FROMINVENTSITEID='' THEN VENDGROUP WHEN FROMINVENTSITEID NOT LIKE '149%' THEN 'BTREC' end [TYPE]  from ax.acxinventtranslocal  a where DATEPHYSICAL BETWEEN '01-aug-2024' and '01-sep-2024' and a.FROMINVENTSITEID like '%v%'group by referenceid , cast(DATEPHYSICAL as date),FROMINVENTSITEID,vendgroup,INVENTSITEID ORDER BY 2"                              
            

            if sql_query:
                df = pd.read_sql(sql_query, engine)

            if df is not None:
                # Get the last row written to the sheet
                startrow = sheet_last_row.get(sheet_name, 0)

                # Write the DataFrame to the sheet starting from the next available row
                df.to_excel(excel_writer, sheet_name=sheet_name, index=False, startrow=startrow)

                # Update the last row written to the sheet
                sheet_last_row[sheet_name] = startrow + len(df) + 1  # Add 1 for spacing between data

    except Exception as e:
        print(f"Error for server {server_name}: {str(e)}")
        continue  # Continue to the next server

# Save the final Excel file
excel_writer._save()

print('Data exported to output_data.xlsx')