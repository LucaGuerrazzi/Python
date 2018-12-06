import pandas as pd
from datetime import datetime
import numpy as np
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import progressbar

#===================================================================
#constant declaration
XLS_COL_RANGE='A:G'
XLS_LOTTO_COL='A'
LOTTO_CESSIONE='LOTTO DI CESSIONE'
MESE_CESSIONE='Mese/Anno di cessione'
NUM_FATTURE='Numero fatture cedute'
IMP_FATTURE='Importo fatture ceduto'
IMP_VER='Importo Versato Al Factor'
RAP_PERC='Rapporto % Ceduto - Versato'
DATA_PAG='Data Pagamento'
DATA_RIF='Data Rif'
SEGMENTO_CLI="SEGMENTO"
SEGMENTO_MIDDLE=1
SEGMENTO_RET_COND=2
SEGMENTO_PDR_RAT=3
UK_MONTH_LIST=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
IT_MONTH_LIST=['GEN','FEB','MAR','APR','MAG','GIU','LUG','AGO','SET','OTT','NOV','DIC']
IT_LNG_MONTH_LIST=['GENNA','FEBBR','MARZO','APRIL','MAGGI','GIUGN','LUGLI','AGOST','SETTE','OTTOB','NOVEM','DICEM']
#===================================================================

def setDataRif(dataString):
    return datetime.strptime(formatDateString(dataString.replace(' ','')),'%m-%y')

def formatDateString(strDate):
    if len(strDate)<4:
        return ''
    elif strDate[0:5].upper() in IT_LNG_MONTH_LIST:
        return str(IT_LNG_MONTH_LIST.index(strDate[0:5].upper()) + 1) + '-' + strDate[-2:]
    elif strDate[0:3].upper() in UK_MONTH_LIST:
        return str(UK_MONTH_LIST.index(strDate[0:3].upper()) + 1) + '-' + strDate[-2:]
    elif strDate[0:3].upper() in IT_MONTH_LIST:
        return str(IT_MONTH_LIST.index(strDate[0:3].upper()) + 1) + '-' + strDate[-2:]
    else:
        return strDate
    
def calcPercPag(df_pag,df_ced,segmento):
    df_pag=df_pag.loc[df_pag[SEGMENTO_CLI]==segmento].T.iloc[1:].reset_index()
    df_ced=df_ced.loc[df_ced[SEGMENTO_CLI]==segmento].drop(columns=SEGMENTO_CLI)

    if (df_pag.empty==False) & (df_ced.empty==False):
        df_check=df_pag.merge(df_ced, left_on='index',right_on=[MESE_CESSIONE],how='inner').drop(columns=MESE_CESSIONE)
        df_check.rename({df_check.columns[1]: 'Importo Pagato', 'index':MESE_CESSIONE}, axis='columns', inplace=True)

        df_check['delta']=df_check[IMP_FATTURE].astype('float64')-df_check['Importo Pagato'].astype('float64')
        df_check['ratio']=df_check['Importo Pagato'].astype('float64')/df_check[IMP_FATTURE].astype('float64')
        df_check=df_check.set_index(MESE_CESSIONE).T
        df_check[DATA_PAG]=df_check.index
        cols = df_check.columns.tolist()
        df_check = df_check[[cols[-1]] + cols[:-1]] 
    else:
        df_check=pd.DataFrame()
        
    return df_check

def isSheetNameOk(sheetName):
    if sheetName.find('-')>0:
        return True
    else:
        return False
    
def getInputAndPrintRotativa():
    Tk().withdraw() 
    xlsName = askopenfilename()
    
    if xlsName=='':
        return 'Selezionare il file da caricare'

    print('Analisi file: ' + xlsName)
    print('Identificazione segmento')
    
    segmento=getSegmento(xlsName)
    if segmento==SEGMENTO_RET_COND:
        print('Segmento identificato ---> Condomini e Retail')
    elif segmento==SEGMENTO_MIDDLE:
        print('Segmento identificato ---> Middle')
    else:
        print('Segmento identificato ---> PDR e RAT')
        
    return printRotativa(xlsName,segmento)

def dfCeduto(df, segmento):
    df=df.loc[df[SEGMENTO_CLI]==segmento].drop(columns=SEGMENTO_CLI)
    df=df.T
    df_new_header=df.iloc[0]
    df=df[1:]
    df.columns=df_new_header
    df[DATA_PAG]="Totale Ceduto"
    cols = df.columns.tolist()
    df = df[[cols[-1]] + cols[:-1]] 
    
    return df

def getSegmento(xlsName):
    xlsStruct = pd.ExcelFile(xlsName)
    
    numElm=len(xlsStruct.sheet_names)
    elmCounter=0
    
    with progressbar.ProgressBar(max_value=numElm) as bar:
        for sheetName in reversed(xlsStruct.sheet_names):
            if isSheetNameOk(sheetName):
                elmCounter+=1
                bar.update(elmCounter)
                df = pd.read_excel(xlsName, sheet_name=sheetName, usecols=XLS_LOTTO_COL, index_col=None, header=None)
                df.dropna(how='all', inplace=True)
                df[0]=df[0].str[0:5]
                df_cond=df.loc[df[0]=='COND_']
                df_pdr_rat=df.loc[(df[0]=='RESRA') | (df[0]=='RESPR') | (df[0]=='PIVAR') | (df[0]=='PIVAP')]
                if df_cond.empty==False:
                    bar.update(numElm)
                    return SEGMENTO_RET_COND
                elif  df_pdr_rat.empty==False:
                    bar.update(numElm)
                    return SEGMENTO_PDR_RAT
        bar.update(numElm)
        return SEGMENTO_MIDDLE
    
def headersEqual(list1, list2):
    if list1==list2:
        return True
    else:
        return False

def printRotativa(xlsName, segmento=SEGMENTO_RET_COND):    
    df_dataset =  pd.DataFrame()
    
    print('Analisi sheet')
    xlsStruct = pd.ExcelFile(xlsName)
    numElm=len(xlsStruct.sheet_names)
    elmCounter=0

    with progressbar.ProgressBar(max_value=numElm) as bar:
        for sheetName in xlsStruct.sheet_names:
            if isSheetNameOk(sheetName):
                elmCounter+=1
                bar.update(elmCounter)
                df = pd.read_excel(xlsName, sheet_name=sheetName, usecols=XLS_COL_RANGE, index_col=None, header=None)
                df.rename({0 : LOTTO_CESSIONE, 1: MESE_CESSIONE,2: NUM_FATTURE,3: IMP_FATTURE, 4: IMP_VER, 5: RAP_PERC, 6: DATA_PAG}, axis='columns', inplace=True)
                df=df.astype(str)
                df=df.loc[(df[DATA_PAG]!='nan') & (df[DATA_PAG]!=DATA_PAG) & (df[DATA_PAG]!='00:00:00') & (df[MESE_CESSIONE]!=MESE_CESSIONE)]             
                df[DATA_RIF]=setDataRif(sheetName)
                df_dataset=df_dataset.append(df)
    bar.update(numElm)
    #===================================================================================
    #costruzione dataframe di sintesi dei lotti
    print('Elaborazione: costruzione dB di sintesi dei lotti')
    df_lotti=df_dataset[[LOTTO_CESSIONE,MESE_CESSIONE,NUM_FATTURE,IMP_FATTURE]].loc[df_dataset[LOTTO_CESSIONE]!='nan' ]
    df_lotti[MESE_CESSIONE]=pd.to_datetime(df_lotti[MESE_CESSIONE], format='%Y-%m-%d %H:%M:%S').dt.date
    df_lotti.sort_values(by=MESE_CESSIONE, inplace=True)
    #===================================================================================  
    #costruzione df ceduto
    print('Elaborazione: costruzione dB ceduto complessivo')
    df_ceduto=df_lotti[[LOTTO_CESSIONE,MESE_CESSIONE,IMP_FATTURE]]
    if segmento==SEGMENTO_MIDDLE:
        df_ceduto[SEGMENTO_CLI]='MIDDLE'
    elif segmento==SEGMENTO_RET_COND:
        df_ceduto[SEGMENTO_CLI]=np.where(df_ceduto[LOTTO_CESSIONE].str.contains('COND'),'CONDOMINIO', 'RETAIL')
    else:
        df_ceduto[SEGMENTO_CLI]=np.where(df_ceduto[LOTTO_CESSIONE].str.contains('RAT'),'RAT', 'PDR')
        
    df_ceduto[IMP_FATTURE]=df_ceduto[IMP_FATTURE].astype(float)
    df_ceduto.drop(columns=LOTTO_CESSIONE, inplace=True)
    df_ceduto=df_ceduto.groupby([SEGMENTO_CLI,MESE_CESSIONE]).sum().reset_index()
    df_ceduto[MESE_CESSIONE]=pd.to_datetime(df_ceduto[MESE_CESSIONE], format='%Y-%m-%d').dt.strftime('%b %Y')
    if segmento==SEGMENTO_MIDDLE:
        df_ceduto_middle=dfCeduto(df_ceduto,'MIDDLE')
    elif segmento==SEGMENTO_RET_COND:
        df_ceduto_retail=dfCeduto(df_ceduto,'RETAIL')
        df_ceduto_condominio=dfCeduto(df_ceduto,'CONDOMINIO')
    else:
        df_ceduto_rat=dfCeduto(df_ceduto,'RAT')
        df_ceduto_pdr=dfCeduto(df_ceduto,'PDR')
    #==================================================================================      
    #costruzione dataframe dettagli
    print('Elaborazione: costruzione dB dettaglio')
    df_dataset.reset_index(inplace=True)
    listCodiciLottoIndex=[indice for indice in df_dataset.index[df_dataset[LOTTO_CESSIONE]!='nan'].tolist()]
    listCodiciLottoIndex.append(df_dataset[DATA_PAG].count()+1)
    numCodiciLotto=len(listCodiciLottoIndex)
    for curIndexCodLotto in range(0,numCodiciLotto-1):
        df_dataset.loc[listCodiciLottoIndex[curIndexCodLotto]:listCodiciLottoIndex[curIndexCodLotto+1]-1,LOTTO_CESSIONE] = df_dataset.loc[listCodiciLottoIndex[curIndexCodLotto],LOTTO_CESSIONE]
        df_check=df_dataset.loc[listCodiciLottoIndex[curIndexCodLotto]:listCodiciLottoIndex[curIndexCodLotto+1]-1,DATA_PAG]
        df_check_ordered=df_check.sort_values()
        if df_check.equals(df_check_ordered)==False:
            print(' WARNING: Possibile inconsistenza date per il lotto ' + df_dataset.loc[listCodiciLottoIndex[curIndexCodLotto],LOTTO_CESSIONE])
    df_dataset=df_dataset[[LOTTO_CESSIONE,IMP_VER,RAP_PERC,DATA_PAG,DATA_RIF]]
    if segmento==SEGMENTO_MIDDLE:
        df_dataset[SEGMENTO_CLI]='MIDDLE'
    elif segmento==SEGMENTO_RET_COND:
        df_dataset[SEGMENTO_CLI]=np.where(df_dataset[LOTTO_CESSIONE].str.contains('COND'),'CONDOMINIO', 'RETAIL')
    else:
        df_dataset[SEGMENTO_CLI]=np.where(df_dataset[LOTTO_CESSIONE].str.contains('RAT'),'RAT', 'PDR')

    df_dataset[DATA_PAG]=pd.to_datetime(df_dataset[DATA_PAG].str[0:4]+df_dataset[DATA_PAG].str[5:7]+df_dataset[DATA_PAG].str[8:10],format="%Y%m%d").dt.date
    df_dataset.sort_values(by=[DATA_RIF,DATA_PAG], inplace=True)
    #===================================================================================
    #costruzione matrice
    print('Elaborazione: costruzione matrice')
    df_per_matr=df_dataset.loc[:,(SEGMENTO_CLI,DATA_RIF,DATA_PAG,IMP_VER)]
    
    df_per_matr[IMP_VER]=df_per_matr[IMP_VER].astype(float)
    df_per_matr=df_per_matr.groupby([SEGMENTO_CLI,DATA_RIF,DATA_PAG]).sum().reset_index()
    df_data_rif=df_per_matr[DATA_RIF].drop_duplicates(keep='last')
    df_transposed_matr=df_per_matr[[SEGMENTO_CLI,DATA_PAG]].drop_duplicates(keep='last')
    for data_riferimento in df_data_rif:
        df_transposed_matr=df_transposed_matr.merge(df_per_matr.loc[df_per_matr[DATA_RIF]==data_riferimento].drop(columns=[DATA_RIF]).rename({IMP_VER : datetime.strftime(data_riferimento,'%b %Y')}, axis='columns'),left_on=[SEGMENTO_CLI,DATA_PAG], right_on=[SEGMENTO_CLI,DATA_PAG],how='outer')
    #===================================================================================
    #check totale ceduto vs. pagamento
    print('Elaborazione: costruzione dati di sintesi')
    df_totale_pagato=df_transposed_matr.groupby([SEGMENTO_CLI]).sum().reset_index()
    if segmento==SEGMENTO_MIDDLE:
        df_sintesi_middle=calcPercPag(df_totale_pagato,df_ceduto,'MIDDLE')
    elif segmento==SEGMENTO_RET_COND:
        df_sintesi_retail=calcPercPag(df_totale_pagato,df_ceduto,'RETAIL')
        df_sintesi_condominio=calcPercPag(df_totale_pagato,df_ceduto,'CONDOMINIO')
    else:
        df_sintesi_rat=calcPercPag(df_totale_pagato,df_ceduto,'RAT')
        df_sintesi_pdr=calcPercPag(df_totale_pagato,df_ceduto,'PDR')
    #===================================================================================
    out_excel = pd.ExcelWriter('rotativa_retail.xlsx', engine='xlsxwriter')
    print('Elaborazione: scrittura excel')
    if segmento==SEGMENTO_MIDDLE:
        df_sintesi_middle.to_excel(out_excel,sheet_name='MIDDLE', index=False)
        df_middle=df_transposed_matr.loc[df_transposed_matr[SEGMENTO_CLI]=='MIDDLE'].drop(columns=[SEGMENTO_CLI]).dropna(axis=1, how='all')
        df_middle.to_excel(out_excel,sheet_name='MIDDLE', header=False, index=False, startrow=5)
        if headersEqual(list(df_ceduto_middle),list(df_middle))==False:
            out_excel.save()        
            return ' WARNING: Ceduto complessivo disallineato con dettaglio'
    elif segmento==SEGMENTO_RET_COND:
        df_sintesi_retail.to_excel(out_excel,sheet_name='RETAIL', index=False)
        df_retail=df_transposed_matr.loc[df_transposed_matr[SEGMENTO_CLI]=='RETAIL'].drop(columns=[SEGMENTO_CLI]).dropna(axis=1, how='all')
        df_retail.to_excel(out_excel,sheet_name='RETAIL', index=False, header=False, startrow=5)   
        df_sintesi_condominio.to_excel(out_excel,sheet_name='CONDOMINI', index=False)   
        df_condominio=df_transposed_matr.loc[df_transposed_matr[SEGMENTO_CLI]=='CONDOMINIO'].drop(columns=[SEGMENTO_CLI]).dropna(axis=1, how='all')
        df_condominio.to_excel(out_excel,sheet_name='CONDOMINI', index=False, header=False, startrow=5)   
        if (headersEqual(list(df_ceduto_retail),list(df_retail))==False) | (headersEqual(list(df_ceduto_condominio),list(df_condominio))==False):
            out_excel.save()
            return 'Ceduto complessivo disallineato con dettaglio'
    elif segmento==SEGMENTO_PDR_RAT:
        df_sintesi_rat.to_excel(out_excel,sheet_name='Rateizzazioni', index=False)
        df_rat=df_transposed_matr.loc[df_transposed_matr[SEGMENTO_CLI]=='RAT'].drop(columns=[SEGMENTO_CLI]).dropna(axis=1, how='all')
        df_rat.to_excel(out_excel,sheet_name='Rateizzazioni', index=False, header=False, startrow=5)   
        df_sintesi_pdr.to_excel(out_excel,sheet_name='Piani di rientro', index=False)   
        df_pdr=df_transposed_matr.loc[df_transposed_matr[SEGMENTO_CLI]=='PDR'].drop(columns=[SEGMENTO_CLI]).dropna(axis=1, how='all')
        df_pdr.to_excel(out_excel,sheet_name='Piani di rientro', index=False, header=False, startrow=5)   
        if (headersEqual(list(df_ceduto_rat),list(df_rat))==False) | (headersEqual(list(df_ceduto_pdr),list(df_pdr))==False):
            out_excel.save()
            return ' WARNING: Ceduto complessivo disallineato con dettaglio'
        
    out_excel.save()        
    
    return 'Elaborazione completata'

if __name__  == "__main__":
    print(getInputAndPrintRotativa())





