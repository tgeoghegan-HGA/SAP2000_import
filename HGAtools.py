import pandas as pd
import numpy as np

def printhello():
    print("hello")

def err_chk(ret = 0, msg = ''):
        if ret != 0: print('Error! ' + msg)
        else: print('ok... ' + msg)

def open_model(sap_obj, SourceModel):
    reternValue = sap_obj.File.OpenFile(SourceModel,)
    err_chk(reternValue, 'Open Model: ' + SourceModel)

def save_model(sap_obj, NewName):
    reternValue = sap_obj.File.Save(NewName)
    err_chk(reternValue, 'Save Model: ' + NewName)

def make_df(keys, TableData, n):
    DF = pd.DataFrame(np.array(TableData).reshape(n, len(keys)).tolist(), columns=keys)
    return DF


def Get_table_E(TableKey, sap_obj):
    print("hello teg table")
    GroupName = 'All'
    TableVersion =  0
    FieldKeyList = ''
    FieldsKeysIncluded = ['']
    NumberRecords =  0
    TableData = ['']
    returnValue =  0
    [TableVersion, FieldsKeysIncluded, NumberRecords, TableData, returnValue] = sap_obj.DatabaseTables.GetTableForEditingArray(TableKey, GroupName, TableVersion, FieldsKeysIncluded, NumberRecords, TableData)
    err_chk(returnValue, 'Get tables for editing')
    print(TableData)
    DF = make_df(FieldsKeysIncluded, TableData, NumberRecords)
    return DF



def mod_table(TableKey, add_content = None, replacing_DF = None):
        GroupName = 'All'
        TableVersion =  0
        FieldsKeysIncluded = ['']
        NumberRecords =  0
        TableData = ['']
        returnValue =  0
        [TableVersion, FieldsKeysIncluded, NumberRecords, TableData, returnValue] = sap_obj.DatabaseTables.GetTableForEditingArray(TableKey,GroupName, TableVersion, FieldsKeysIncluded,NumberRecords, TableData)
        err_chk(returnValue, 'GetTableForEditingArray'+ TableKey)
        keys = FieldsKeysIncluded
        n = NumberRecords
        DF = make_df(keys, TableData, n)
        if add_content != None:
            if len(np.array(add_content).shape) == 1:
                DF = pd.concat([DF, make_df(keys, add_content, n= 1)]).reset_index(drop = True)
            else:
                for row in add_content:
                    DF = pd.concat([DF, make_df(keys, row, n= 1)]).reset_index(drop = True)
        elif type(replacing_DF) != type(None):
            DF = replacing_DF
        TableData = tuple(np.array(DF).flatten())
        print(TableData)
        [TableVersion, NumberRecords, TableData, returnValue] = sap_obj.DatabaseTables.SetTableForEditingArray(TableKey,TableVersion, FieldsKeysIncluded,NumberRecords, TableData)
        err_chk(returnValue, 'SetTableForEditingArray for: ' + TableKey)
        return DF

def Apply_edits():
    # Apply edits    
    FillImportLog = 1
    NumFatalErrors =  0
    NumErrorMsgs =  0
    NumWarnMsgs =  0
    NumInfoMsgs =  0
    ImportLog = ''

    [NumFatalErrors, NumErrorMsgs, NumWarnMsgs, NumInfoMsgs, ImportLog, returnValue] = sap_obj.DatabaseTables.ApplyEditedTables(FillImportLog,NumFatalErrors, NumErrorMsgs, NumWarnMsgs,NumInfoMsgs, ImportLog)
    err_chk(NumErrorMsgs, 'ApplyEditedTables')
    print(ImportLog)
    
def get_ID_DF(elm):
    DF = Get_table(elm + ' Object Connectivity', sap_obj)
    if elm == 'Point': DF['ID'] = DF['X'] + DF['Y'] + DF['Z']
    if elm == 'Beam' or elm == 'Column' : DF['ID'] = DF['UniquePtI'] + DF['UniquePtJ']
    if elm == 'Wall' or elm == 'Floor': 
        cols = DF.columns[DF.columns.str.contains('UniquePt')]
        DF['ID'] = DF[cols].sum(axis=1)
    return DF

def UpdateName(Name, NewName, ObjType):
    if ObjType == 'Area' : instance = sap_obj.AreaObj
    if ObjType == 'Joint' : instance = sap_obj.PointObj
    if ObjType == 'Frame' : instance = sap_obj.FrameObj
    returnValue =  0  
    err_chk(instance.ChangeName(Name,NewName), 'Updating Name of ' + Name  + ' to ' + NewName)

def float_it(DF, float_cols = False, cols = None):
    # DF : dataframe to be converted, obviously
    # float_cols: when true, returns a list of converted columns
    # cols : list of column to be converted to float
    if cols == None: cols = DF.columns
    for col in cols:
        try: DF[col] = DF[col].astype(float)
        except: pass
    if float_cols: return [DF, DF.dtypes[DF.dtypes== float].keys().tolist()]
    else: return DF

def Get_table(TableKey, sap_obj):
    GroupName = 'All'
    TableVersion =  0
    FieldKeyList = ''
    FieldsKeysIncluded = ['']
    NumberRecords =  0
    TableData = ['']
    returnValue =  0
    [GroupName, TableVersion, FieldsKeysIncluded, NumberRecords, TableData, returnValue] = sap_obj.DatabaseTables.GetTableForDisplayArray(TableKey, FieldKeyList, GroupName, TableVersion, FieldsKeysIncluded, NumberRecords, TableData)
    # print(TableData)
    DF = make_df(FieldsKeysIncluded, TableData, NumberRecords)
    return DF