import pandas as pd
import os
import numpy as np

pvinfile="\pv-sheet.xlsx"
sapinfile = "\sap-sheet.xlsx"
sapOutfile= "\sap-results.xlsx"
folder=os.path.dirname(__file__)


def FileExists(file,isfolder=False)-> bool:
   
    if isfolder:
        return os.path.isdir(file)
    else:
        return os.path.isfile(file)

def getdatafile(datafile,folder)-> pd.DataFrame:
    dataframe=pd.DataFrame()
    ipfile=folder+datafile
    try:
        if FileExists(ipfile):
            dataframe=pd.read_excel(ipfile,sheet_name='Sheet1')
        else:
            print(f'Cannot find input file {ipfile}')
    except Exception as err:  
        print(f'Failed getting input file {datafile}')
        print(err)
    return dataframe

pvdf=getdatafile(pvinfile,folder)
sapdf = getdatafile(sapinfile,folder)

def createresultsdf(sapf:pd.DataFrame,pvdf:pd.DataFrame)-> pd.DataFrame:
    ##This method sizes an empty df with 'NA' the same size as the passed df and uses the same columns
    copyDF=pvdf.copy()
    copyDF.rename(columns = {'TAG':'PVTAG'}, inplace = True) # best not to have two columns with the same name
    copyDF[:]='NA'
    return pd.concat([sapf,copyDF],axis=1) ##merge the two df together tho get one new one that includes a default NA for all cells


def assignmatchrow(index:int,sapdf:pd.DataFrame,matchedrow:pd.DataFrame):
    totalCols=len(sapdf.columns)
    offset=totalCols-len(matchedrow.columns)
    try:
        for i in range(len(matchedrow.columns)):
            sapdf.iloc[index,i+offset]=matchedrow.iloc[0,i]
    except Exception as err:
        print(f'Error in updatedf method')
        print(err)   


def excelsave(df: pd.DataFrame, file: str, sheet: str = 'Sheet1') -> bool:
    try:
        with pd.ExcelWriter(file) as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)
        return True
    except FileNotFoundError as e:
        print(f"Error: file '{file}' not found.")
        return False
    except PermissionError as e:
        print(f"Error: permission denied for file '{file}'.")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False



def comparedf(sapdf:pd.DataFrame, pvfile:pd.DataFrame)-> pd.DataFrame:
    try:
        for index,row in sapdf.iterrows():
            lookup=pvfile.loc[pvfile['TAG'] == row['TAG']]  #this generates a dataframe containing the matched PV row
            if not lookup.empty:
                print(f'{row["TAG"]} matched in PVfile')
                assignmatchrow(index,sapdf,lookup)
        return sapdf
        
    except Exception as err:
        print('Error in cmparedf')
        print(err)

        
        
newSapDf=createresultsdf(sapdf,pvdf) ## generate a new df with all cells=NA and right merge it to the sap df
comparedf(newSapDf, pvdf) # pass in the new resized sap DF instead of the original one

excelsave(newSapDf,folder+sapOutfile)