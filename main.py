import pandas as pd
import os
import numpy as np

pvinfile="\pv-sheet.xlsx"
sapinfile = "\sap-sheet.xlsx"
sapOutfile= "\sap-results.xlsx"
folder=os.path.dirname(__file__)


def FileExists(file,isfolder=False):
        if isfolder:
            return os.path.isdir(file)
        else:
            return os.path.isfile(file)

def getdatafile(datafile,folder):
    dataframe=pd.DataFrame()
    ipfile=folder+datafile
    try:
        if FileExists(ipfile):
            dataframe=pd.read_excel(ipfile,sheet_name='Sheet1')
        else:
            print(f'Cannot find input file {ipfile}')
        #dataframe=pd.read_csv(datafile, na_values=[' '],keep_default_na=False,dtype=str, sep=";")

    except Exception as err:  
        print(f'Failed getting input file {datafile}')
        print(err)
    return dataframe

pvdf=getdatafile(pvinfile,folder)
sapdf = getdatafile(sapinfile,folder)

def createresultsdf(sapf:pd.DataFrame,pvdf:pd.DataFrame):
    ##This method sizes an empty df with 'NA' the same size as the passed df and uses the same columns
    copyDF=pvdf.copy()
    copyDF.rename(columns = {'TAG':'PVTAG'}, inplace = True) # best not to have two columns with the same name
    copyDF[:]='NA' const periodType = "MonthPeriod";
tau.mashups
    .addDependency('tau/api/configurable-controls/controls/v1')
    .addMashup(( controlsApiV1 ) => {
        controlsApiV1.addConfiguration('dropdown', {
          id: 'dd_conf_v2_available_user',
          name: 'Available User',
          label: 'Available User',
          supportedEntityTypes: ['workallocation'],
          requiredEntityFields: [
            'connectedUser:connectedUser==null?null:{id:connectedUser.id,name:connectedUser.fullName,resourceType:connectedUser.resourceType}',
            `periods: demands.select(${periodType}.id)`,
            'percentage: [Percentage]',
            'type: CustomValues["Allocation Type"]'
          ],
          sampleData: {
            connectedUser: { name: 'John Smith' , id: 5}
          },
          entityClickBehavior: 'openEntity',
          field: 'connectedUser',
          dropdownItemsSource: {
            type: 'custom',
            getItems: async options => {
                const args = options.entity;
                let type = args["type"];
                if (options.renderingLocation === 'quickAdd') {
                    type = document.querySelector('.quick-add form.WorkAllocation.tau-active [data-fieldname="Allocation Type"]').value;
                }
                const periods = args.periods||[];
                const percentage = args.percentage||0;
                if( !type || type === 'User'){
                    const response = await fetch(`/api/v2/user?select={id, fullname as name}&where=(Availabilities.Where(${periodType}.Id in ${JSON.stringify(periods)}).Count()==${periods.length} and Availabilities.Where(${periodType}.Id in ${JSON.stringify(periods)}).Where((AvailableCapacity)<${percentage}).Count()==0)&take=999`);
                    const {items} = await response.json();
                    return items;
                }
                return []
            },
          }
        });
        controlsApiV1.addConfiguration('dropdown', {
          id: 'dd_conf_v2_connected_user',
          name: 'Connected User',
          label: 'Connected User',
          supportedEntityTypes: ['workallocation'],
          requiredEntityFields: [
            'connectedUser:connectedUser==null?null:{id:connectedUser.id,name:connectedUser.fullName,resourceType:connectedUser.resourceType}',
            'type: CustomValues["Allocation Type"]'
          ],
          sampleData: {
            connectedUser: { name: 'John Smith' , id: 5}
          },
          entityClickBehavior: 'openEntity',
          field: 'connectedUser',
          dropdownItemsSource: {
            type: 'custom',
            getItems: async options => {
                const args = options.entity;
                let type = args["type"];
                if (options.renderingLocation === 'quickAdd') {
                    type = document.querySelector('.quick-add form.WorkAllocation.tau-active [data-fieldname="Allocation Type"]').value;
                }
                const periods = args.periods||[];
                const percentage = args.percentage||0;
                if( !type || type === 'User'){
                    const response = await fetch(`/api/v2/user?select={id, fullname as name}&take=1000`);
                    const {items} = await response.json();
                    return items;
                }
                return []
            },
          }
        });
    })
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


def excelsave(df:pd.DataFrame, file:str,sheet:str='Sheet1'):
    try:
        df.to_excel(file,sheet_name=sheet,index = False)
    except Exception as err:
        print(f"Error in compare_spo_to_sap \n{err}")


def comparedf(sapdf:pd.DataFrame, pvfile:pd.DataFrame):
    try:

        for index,row in sapdf.iterrows():
            lookup=pvfile.loc[pvfile['TAG'] == row['TAG']]  #this generates a dataframe containing the matched PV row
            if lookup.empty==False:
                print(f'{row["TAG"]} matched in PVfile')
                assignmatchrow(index,sapdf,lookup)
        return sapdf
        
    except Exception as err:
        print('Error in cmparedf')
        print(err)

        
        
newSapDf=createresultsdf(sapdf,pvdf) ## generate a new df with all cells=NA and right merge it to the sap df
comparedf(newSapDf, pvdf) # pass in the new resized sap DF instead of the original one

excelsave(newSapDf,folder+sapOutfile)