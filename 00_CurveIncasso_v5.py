import pandas as pd
import numpy as np

#===================================================================
#constant declaration
DB_FILE = 'Db.xlsx'
DB_PATH = ''

SHEET_CI= 'DB Curve Incasso'
SHEET_FATT= 'DB Fatturato'

XLS_FATT_HEADER_ROW=5
XLS_CI_HEADER_ROW=5
XLS_FATT_COL_RANGE='B:BI'
XLS_CI_COL_RANGE='B:U'

FATT_TO_GET='Fatturato in scadenza'
FATT_TIPOLOGIA='Tipologia'

COL_DATA_SCAD='Data Scadenza'
COL_PROD='Prodotto'
COL_SOTTO_TIP='Sotto-tipologia'
COL_MACRO_SEG='Macro Segmento Cliente'
COL_SEG='Segmento Cliente'
COL_SIST_FATT='Sistema Fatturazione'
COL_MOD_PAG='Modalità Pagamento'
COL_VAR='variable'
COL_MERC='Mercato'

OK_VALUE='OK'

SUFF_FATT='_fatt'
SUFF_INC='_incasso'
OK_CI='OK_CI'
OK_FATT='OK_FATT'

MONTH_LIST=['JAN','GEN','FEB','MAR','APR','MAY','MAG','JUN','GIU','JUL','LUG','AUG','AGO','SEP','SET','OCT','OTT','NOV','DEC','DIC']
CAT_COL=[COL_SOTTO_TIP,COL_PROD,COL_MACRO_SEG,COL_SEG,COL_MERC,COL_SIST_FATT,COL_MOD_PAG]
CAT_COL_E_SCAD=[COL_SOTTO_TIP,COL_PROD,COL_MACRO_SEG,COL_SEG,COL_MERC,COL_SIST_FATT,COL_MOD_PAG,COL_DATA_SCAD]
CAT_COL_E_VAR=[COL_SOTTO_TIP,COL_PROD,COL_MACRO_SEG,COL_SEG,COL_MERC,COL_SIST_FATT,COL_MOD_PAG,COL_VAR]
#===================================================================

def getMissingCombinations(myDataframe,colName1 = "",colName2=""):
    if colName1 != "": 
        myDataframe=myDataframe.loc[myDataframe[colName1]!=OK_VALUE]
    if colName2 != "": 
        myDataframe=myDataframe.loc[myDataframe[colName2]!=OK_VALUE]
        
    myDataframe=myDataframe[CAT_COL].drop_duplicates(keep='last')
    
    return myDataframe
    
def cleanDataframe(myDataframe,orig_sheet):
    if orig_sheet==SHEET_FATT:
        myDataframe=myDataframe.loc[myDataframe[FATT_TIPOLOGIA]==FATT_TO_GET].drop(columns=[FATT_TIPOLOGIA])
        headerList=list(myDataframe)
        for lblLabel in headerList:
            if isMonth(lblLabel[0:3]):
                myDataframe.rename({lblLabel : lblLabel[0:8] + SUFF_FATT}, axis='columns', inplace=True)
                myDataframe[lblLabel[0:8] + SUFF_INC]=0
            elif lblLabel[0:4].isdigit():   #elimina le colonne con header del tipo 2018 PLAN
                myDataframe=myDataframe.drop(columns=[lblLabel])
        myDataframe.rename({'Macro Segmento\nCliente' : COL_MACRO_SEG,'Segmento\nCliente' : COL_SEG,'Sistema \nFatturazione' : COL_SIST_FATT,'Modalità \nPagamento' : COL_MOD_PAG}, axis='columns', inplace=True)
    elif orig_sheet==SHEET_CI:
        myDataframe.rename({'Macro Segmento\nCliente' : COL_MACRO_SEG,'Sistema \nFatturazione' : COL_SIST_FATT,'Modalità \nPagamento' : COL_MOD_PAG}, axis='columns', inplace=True)

    return myDataframe

def isMonth(strMese):
    if strMese.upper in MONTH_LIST:
        return True
    else:
        return False
    
def shift_month(month, shifting):
    list_month=MONTH_LIST
    column_number=list_month.index(month[0:3])+shifting
    anno=int(month[4:8])
    delta_anno = column_number // 12
    if column_number>11:
        column_number=column_number % 12

    return list_month[column_number] + ' ' + str(anno + delta_anno)

def calcIncassi():
    df = pd.read_excel(DB_PATH + DB_FILE, sheet_name=SHEET_FATT, usecols=XLS_FATT_COL_RANGE, index_col=None, header=XLS_FATT_HEADER_ROW)
    df_fatt=cleanDataframe(df,SHEET_FATT)
   
    df = pd.read_excel(DB_PATH + DB_FILE, sheet_name=SHEET_CI, usecols=XLS_CI_COL_RANGE, index_col=None, header=XLS_CI_HEADER_ROW)
    df_ci_cleaned=cleanDataframe(df,SHEET_CI)
    
    #=====================================================================================
    #realizza la transposizione parziale della tabela delle curve di incasso
    df_ci=pd.melt(df_ci_cleaned, id_vars=CAT_COL_E_SCAD)
    df_ci_data_scad=df_ci[COL_DATA_SCAD].drop_duplicates(keep='last')
    df_ci_transposed_seed=df_ci[CAT_COL_E_VAR].drop_duplicates(keep='last')
    for values in df_ci_data_scad:
        df_ci_transposed_seed=df_ci_transposed_seed.merge(df_ci.loc[df_ci[COL_DATA_SCAD]==values].drop(columns=[COL_DATA_SCAD]).rename({'value' : values}, axis='columns'),left_on=CAT_COL_E_VAR, right_on=CAT_COL_E_VAR,how='inner')
    df_ci_transposed_seed.sort_values(by=CAT_COL_E_VAR,inplace=True)
    #=====================================================================================

    df_ci_transposed_seed[OK_CI]=OK_VALUE       
    df_fatt[OK_FATT]=OK_VALUE
    df_dataset = pd.merge(df_fatt, df_ci_transposed_seed, how="outer", on=CAT_COL)
    
    df_no_fatt=getMissingCombinations(df_dataset,colName1=OK_FATT)
    df_no_ci=getMissingCombinations(df_dataset,colName1=OK_CI)
    
    df_dataset=df_dataset.loc[(df_dataset[OK_FATT]==OK_VALUE) & (df_dataset[OK_CI]==OK_VALUE)].drop(columns=[OK_FATT]).drop(columns=[OK_CI])
    
    df_ci=df_ci_transposed_seed.drop(columns=[OK_CI])
    
    df_incasso = df_dataset
    df_mese_x=df_dataset[COL_VAR].drop_duplicates(keep='last')
    for mese_x in df_mese_x:
        for values in df_ci_data_scad:
            df_incasso.loc[(df_incasso[COL_VAR]==mese_x),shift_month(values[0:8], int(mese_x.replace('MESE ',''))) + SUFF_INC] = df_incasso.loc[(df_incasso[COL_VAR]==mese_x), values + SUFF_FATT] * df_incasso[values]
            
    df_incasso.drop(columns=[colName for colName in list(df_incasso) if (SUFF_FATT in colName) or (isMonth(colName[0:3]) and len(colName)==8)], inplace=True)
    df_incasso.fillna('-',inplace=True)
    df_incasso_agg=df_incasso.groupby(CAT_COL).sum().reset_index()
    df_incasso_agg[COL_VAR]='MESE ALL'
    df_incasso=df_incasso.append(df_incasso_agg)
    df_incasso[COL_VAR]=np.where(df_incasso[COL_VAR].apply(len)==6,df_incasso[COL_VAR].str[0:5] + '0' + df_incasso[COL_VAR].str[5],df_incasso[COL_VAR])
    df_incasso.sort_values(by=CAT_COL_E_VAR, inplace=True)
   
    out_excel = pd.ExcelWriter('curve_incasso.xlsx', engine='xlsxwriter')
    df_incasso.to_excel(out_excel,sheet_name='incasso', startrow=1, header=False, index=False)
    
    workbook  = out_excel.book
    worksheet = out_excel.sheets['incasso']
    header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#ff8d03', 'border': 1})
    for col_num, value in enumerate(df_incasso.columns.values):
        worksheet.write(0, col_num + 1, value, header_format)                                 
    df_no_fatt.to_excel(out_excel,sheet_name='si ci - no fatt', index=False)
    df_no_ci.to_excel(out_excel,sheet_name='no ci - si fatt', index=False)
    out_excel.save()    

if __name__  == "__main__":
    calcIncassi()