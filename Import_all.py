# -*- coding: utf-8 -*-
"""
Created on Sun Dec 30 14:51:02 2018

@author: mikha
"""

import pyodbc
import pandas as pd
import xlrd

server = 'DESKTOP-HP138CN\SQLEXPRESS' #server name
database = 'Risks' 

connection = pyodbc.connect('Driver={SQL Server};SERVER=' + server 
                            + ';DATABASE='+ database
                            + ';Trusted_Connection=yes') #create new connection with server
cursor = connection.cursor() #create special object for queries			

cursor.execute("""
CREATE TABLE [dbo].[Instrs](
    ID nvarchar(1000),
    EfirCode nvarchar(1000), 
    ShortNameRus nvarchar(1000),
    FullNameRus nvarchar(1000),
    ISIN nvarchar(1000),
    EfirCFI nvarchar(1000), 
    CFIName nvarchar(1000),
    Exchange nvarchar(1000),
    ExchTicker nvarchar(1000), 
    ExchSymbol nvarchar(1000),
    EmitentCode nvarchar(1000),
    EmitentName nvarchar(1000),
    MarketSector nvarchar(1000),
    LotSize nvarchar(1000),
    ExpDate nvarchar(1000),
    Currency nvarchar(1000),
    Visible nvarchar(1000),
    RegNum nvarchar(1000)
)""")
connection.commit()

data = pd.read_excel('bd.xlsx') #BOND_DESCRIPTION(INSTRS)
data = data.rename(columns={'ID': 'ID',
                            'EfirCode': 'EfirCode',
                            'ShortNameRus': 'ShortNameRus',
                            'FullNameRus': 'FullNameRus',
                            'ISIN': 'ISIN',
                            'EfirCFI': 'EfirCFI',
                            'CFIName': 'CFIName',
                            'Exchange': 'Exchange',
                            'ExchTicker': 'ExchTicker',
                            'ExchSymbol': 'ExchSymbol',
                            'EmitentCode': 'EmitentCode',
                            'EmitentName': 'EmitentName',
                            'MarketSector': 'MarketSector',
                            'LotSize': 'LotSize',
                            'ExpDate': 'ExpDate',
                            'Currency': 'Currency',
                            'Visible': 'Visible',
                            'RegNum': 'RegNum'}) # rename columns

# export
data.to_excel('bd.xlsx', index=False)

# Open the workbook and define the worksheet
book = xlrd.open_workbook('bd.xlsx')
sheet = book.sheet_by_name('Sheet1')

query1 = ("""
INSERT INTO [Risks].[dbo].[Instrs](
    ID,
    EfirCode, 
    ShortNameRus,
    FullNameRus,
    ISIN,
    EfirCFI, 
    CFIName,
    Exchange,
    ExchTicker, 
    ExchSymbol,
    EmitentCode,
    EmitentName,
    MarketSector,
    LotSize,
    ExpDate,
    Currency,
    Visible,
    RegNum) 
VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
)""")

for r in range(1, sheet.nrows): #read loop for 2nd row
    ID = sheet.cell(r,0).value
    EfirCode = sheet.cell(r,1).value
    ShortNameRus = sheet.cell(r,2).value
    FullNameRus = sheet.cell(r,3).value
    ISIN = sheet.cell(r,4).value
    EfirCFI = sheet.cell(r,5).value
    CFIName = sheet.cell(r,6).value
    Exchange = sheet.cell(r,7).value
    ExchTicker = sheet.cell(r,8).value
    ExchSymbol = sheet.cell(r,9).value
    EmitentCode = sheet.cell(r,10).value
    EmitentName = sheet.cell(r,11).value
    MarketSector = sheet.cell(r,12).value
    LotSize = sheet.cell(r,13).value
    ExpDate = sheet.cell(r,14).value
    Currency = sheet.cell(r,15).value
    Visible = sheet.cell(r,16).value
    RegNum = sheet.cell(r,17).value
    
    # Assign values from each row
    values = (ID,
              EfirCode, 
              ShortNameRus,
              FullNameRus,
              ISIN,
              EfirCFI, 
              CFIName,
              Exchange,
              ExchTicker, 
              ExchSymbol,
              EmitentCode,
              EmitentName,
              MarketSector,
              LotSize,
              ExpDate,
              Currency,
              Visible,
              RegNum)
    cursor.execute(query1, values)
connection.commit()

cursor.execute("""
CREATE TABLE [dbo].[Bond](
    ISIN_RegCode_NRDCode nvarchar(1000),
    FinToolType nvarchar(1000),
    Sec_Type_ID	nvarchar(1000),
    SecTypeNameRus_NRD nvarchar(1000),
    SecTypeNameEng_NRD nvarchar(1000),
    Sec_Type_BR_Code nvarchar(1000),
    SecTypeNameBR_NRD nvarchar(1000),
    CFI	nvarchar(1000),
    SecurityType nvarchar(1000),
    SecurityKind nvarchar(1000),
    Sec_Form_ID	nvarchar(1000),
    SecFormNameRus_NRD nvarchar(1000),
    SecFormNameEng_NRD nvarchar(1000),
    CouponType nvarchar(1000),
    Rate_Type_ID nvarchar(1000),
    RateTypeNameRus_NRD nvarchar(1000),
    RateTypeNameEng_NRD	nvarchar(1000),
    CouponTypeName_NRD nvarchar(1000),
    HaveOffer nvarchar(1000),
    AmortisedMty nvarchar(1000),
    MaturityGroup nvarchar(1000),
    IsConvertible nvarchar(1000),
    NickName nvarchar(1000),
    FullName nvarchar(1000),
    FullName_NRD nvarchar(1000),
    ISINCode nvarchar(1000),
    ISIN144A nvarchar(1000),
    NRDCode nvarchar(1000),
    RegCode nvarchar(1000),
    RegCode_M nvarchar(1000),
    RegCode_NRD nvarchar(1000),
    IssueNumber nvarchar(1000),
    BondSeries nvarchar(1000),
    FinToolID nvarchar(1000),
    ProgFinToolID nvarchar(1000),
    Status nvarchar(1000),
    Sec_State_ID nvarchar(1000),
    SecStateRus_NRD nvarchar(1000),
    SecStateEng_NRD nvarchar(1000),
    HaveDefault nvarchar(1000),
    GuaranteeType nvarchar(1000),
    GuaranteeAmount nvarchar(1000),
    BorrowerName nvarchar(1000),
    BorrowerOKPO nvarchar(1000),
    BorrowerSector nvarchar(1000),
    BorrowerUID nvarchar(1000),
    IssuerName nvarchar(1000),
    IssuerName_NRD nvarchar(1000),
    IssuerOKPO nvarchar(1000),
    IssuerUID nvarchar(1000),
    NumGuarantors nvarchar(1000),
    Acc_Open_Date_NRD nvarchar(1000),
    FacialAcc_NRD nvarchar(1000),
    RegistrarAccTypeDate_NRD nvarchar(1000),
    RegistrarAccType_NRD nvarchar(1000),
    Registrar_NRD nvarchar(1000),
    PrivateDist nvarchar(1000),
    Placement_Type_ID nvarchar(1000),
    PlacementType_NRD nvarchar(1000),
    RegDate nvarchar(1000),
    RegDate_M nvarchar(1000),
    RegDate_NRD nvarchar(1000),
    RegOrg nvarchar(1000),
    RegOrg_NRD nvarchar(1000),
    BegDistDate nvarchar(1000),
    BegDistDate_M nvarchar(1000),
    BegDistDate_NRD nvarchar(1000),
    EndDistDate nvarchar(1000),
    EndDistDate_M nvarchar(1000),
    EndDistDate_NRD nvarchar(1000),
    RegDistDate nvarchar(1000),
    RegDistDate_M nvarchar(1000),
    RegDistDate_NRD nvarchar(1000),
    RpState_NRD nvarchar(1000),
    RpRegOrg_NRD nvarchar(1000),
    PlacePrice_NRD nvarchar(1000),
    SumIssueVal nvarchar(1000),
    SumIssueVol nvarchar(1000),
    SumIssueVol_M nvarchar(1000),
    SumIssueVol_NRD nvarchar(1000),
    SumMarketVal nvarchar(1000),
    SumMarketVol nvarchar(1000),
    SumMarketVol_M nvarchar(1000),
    SumMarketVol_NRD nvarchar(1000),
    FirstCouponDate_NRD nvarchar(1000),
    BegMtyDate nvarchar(1000),
    PlanMtyDate_NRD nvarchar(1000),
    EndMtyDate nvarchar(1000),
    EndMtyDate_M nvarchar(1000),
    EndMtyDate_NRD nvarchar(1000),
    DaysAll nvarchar(1000),
    DaysAll_M nvarchar(1000),
    DaysAll_NRD nvarchar(1000)
)""")
connection.commit()

data = pd.read_excel('bd2.xlsx') #BOND_DESCRIPTION(BOND_DESCRIPTION)
data = data.rename(columns={'ISIN_RegCode_NRDCode': 'ISIN_RegCode_NRDCode',
                            'EfirCode': 'EfirCode',
                            'FinToolType': 'FinToolType',
                            'Sec_Type_ID': 'Sec_Type_ID',
                            'SecTypeNameRus_NRD': 'SecTypeNameRus_NRD',
                            'SecTypeNameEng_NRD': 'SecTypeNameEng_NRD',
                            'Sec_Type_BR_Code': 'Sec_Type_BR_Code',
                            'SecTypeNameBR_NRD': 'SecTypeNameBR_NRD',
                            'SecTypeNameBR_NRD': 'SecTypeNameBR_NRD',
                            'CFI': 'CFI',
                            'SecurityType': 'SecurityType',
                            'SecurityKind': 'SecurityKind',
                            'Sec_Form_ID': 'Sec_Form_ID',
                            'SecFormNameRus_NRD': 'SecFormNameRus_NRD',
                            'SecFormNameEng_NRD': 'SecFormNameEng_NRD',
                            'CouponType': 'CouponType',
                            'Rate_Type_ID': 'Rate_Type_ID',
                            'RateTypeNameRus_NRD': 'RateTypeNameRus_NRD',
                            'RateTypeNameEng_NRD': 'RateTypeNameEng_NRD',
                            'CouponTypeName_NRD': 'CouponTypeName_NRD',
                            'HaveOffer': 'HaveOffer',
                            'AmortisedMty': 'AmortisedMty',
                            'MaturityGroup': 'MaturityGroup',
                            'IsConvertible': 'IsConvertible',
                            'NickName': 'NickName',
                            'FullName': 'FullName',
                            'FullName_NRD': 'FullName_NRD',
                            'ISINCode': 'ISINCode',
                            'ISIN144A': 'ISIN144A',
                            'NRDCode': 'NRDCode',
                            'RegCode': 'RegCode',
                            'RegCode_M': 'RegCode_M',
                            'RegCode_NRD': 'RegCode_NRD',
                            'IssueNumber': 'IssueNumber',
                            'BondSeries': 'BondSeries',
                            'FinToolID': 'FinToolID',
                            'ProgFinToolID': 'ProgFinToolID',
                            'Status': 'Status',
                            'Sec_State_ID': 'Sec_State_ID',
                            'SecStateRus_NRD': 'SecStateRus_NRD',
                            'SecStateEng_NRD': 'SecStateEng_NRD',
                            'HaveDefault': 'HaveDefault',
                            'GuaranteeType': 'GuaranteeType',
                            'GuaranteeAmount': 'GuaranteeAmount',
                            'BorrowerName': 'BorrowerName',
                            'BorrowerOKPO': 'BorrowerOKPO',
                            'BorrowerSector': 'BorrowerSector',
                            'BorrowerUID': 'BorrowerUID',
                            'IssuerName': ' IssuerName',
                            'IssuerName_NRD': 'IssuerName_NRD',
                            'IssuerOKPO': 'IssuerOKPO',
                            'IssuerUID': 'IssuerUID',
                            'NumGuarantors': 'NumGuarantors',
                            'Acc_Open_Date_NRD': 'Acc_Open_Date_NRD',
                            'FacialAcc_NRD': 'FacialAcc_NRD',
                            'RegistrarAccTypeDate_NRD': 'RegistrarAccTypeDate_NRD',
                            'RegistrarAccType_NRD': 'RegistrarAccType_NRD',
                            'Registrar_NRD': 'Registrar_NRD',
                            'PrivateDist': 'PrivateDist',
                            'Placement_Type_ID': 'Placement_Type_ID',
                            'PlacementType_NRD': 'PlacementType_NRD',
                            'RegDate': 'RegDate',
                            'RegDate_M': 'RegDate_M',
                            'RegDate_NRD': 'RegDate_NRD',
                            'RegOrg': 'RegOrg',
                            'RegOrg_NRD': 'RegOrg_NRD',
                            'BegDistDate': 'BegDistDate',
                            'BegDistDate_M': 'BegDistDate_M',
                            'BegDistDate_NRD': ' BegDistDate_NRD',
                            'EndDistDate': 'EndDistDate',
                            'EndDistDate_M': 'EndDistDate_M',
                            'EndDistDate_NRD': 'EndDistDate_NRD',
                            'RegDistDate': ' RegDistDate',
                            'RegDistDate_M': 'RegDistDate_M',
                            'RegDistDate_NRD': 'RegDistDate_NRD',
                            'RpState_NRD': 'RpState_NRD',
                            'RpRegOrg_NRD': 'RpRegOrg_NRD',
                            'PlacePrice_NRD': 'PlacePrice_NRD',
                            'SumIssueVal': 'SumIssueVal',
                            'SumIssueVol': 'SumIssueVol',
                            'SumIssueVol_M': 'SumIssueVol_M',
                            'SumIssueVol_NRD': 'SumIssueVol_NRD',
                            'SumMarketVal': 'SumMarketVal',
                            'SumMarketVol': 'SumMarketVol',
                            'SumMarketVol_M': 'SumMarketVol_M',
                            'SumMarketVol_NRD': 'SumMarketVol_NRD',
                            'FirstCouponDate_NRD': 'FirstCouponDate_NRD',
                            'BegMtyDate': 'BegMtyDate',
                            'PlanMtyDate_NRD': 'PlanMtyDate_NRD',
                            'EndMtyDate': 'EndMtyDate',
                            'EndMtyDate_M': 'EndMtyDate_M',
                            'EndMtyDate_NRD': 'EndMtyDate_NRD',
                            'DaysAll': 'DaysAll',
                            'DaysAll_M': 'DaysAll_M',
                            'DaysAll_NRD': 'DaysAll_NRD'}) # rename columns

# export
data.to_excel('bd2.xlsx', index=False)

# Open the workbook and define the worksheet
book = xlrd.open_workbook('bd2.xlsx')
sheet = book.sheet_by_name('Sheet1')

query2 = ("""
INSERT INTO [Risks].[dbo].[Bond](
    ISIN_RegCode_NRDCode,
    FinToolType,
    Sec_Type_ID,
    SecTypeNameRus_NRD,
    SecTypeNameEng_NRD,
    Sec_Type_BR_Code,
    SecTypeNameBR_NRD,
    CFI,
    SecurityType,
    SecurityKind,
    Sec_Form_ID,
    SecFormNameRus_NRD,
    SecFormNameEng_NRD,
    CouponType,
    Rate_Type_ID,
    RateTypeNameRus_NRD,
    RateTypeNameEng_NRD,
    CouponTypeName_NRD,
    HaveOffer,
    AmortisedMty,
    MaturityGroup,
    IsConvertible,
    NickName,
    FullName,
    FullName_NRD,
    ISINCode,
    ISIN144A,
    NRDCode,
    RegCode,
    RegCode_M,
    RegCode_NRD,
    IssueNumber,
    BondSeries,
    FinToolID,
    ProgFinToolID,
    Status,
    Sec_State_ID,
    SecStateRus_NRD,
    SecStateEng_NRD,
    HaveDefault,
    GuaranteeType,
    GuaranteeAmount,
    BorrowerName,
    BorrowerOKPO,
    BorrowerSector,
    BorrowerUID,
    IssuerName,
    IssuerName_NRD,
    IssuerOKPO,
    IssuerUID,
    NumGuarantors,
    Acc_Open_Date_NRD,
    FacialAcc_NRD,
    RegistrarAccTypeDate_NRD,
    RegistrarAccType_NRD,
    Registrar_NRD,
    PrivateDist,
    Placement_Type_ID,
    PlacementType_NRD,
    RegDate,
    RegDate_M,
    RegDate_NRD,
    RegOrg,
    RegOrg_NRD, 
    BegDistDate,
    BegDistDate_M,
    BegDistDate_NRD,
    EndDistDate,
    EndDistDate_M,
    EndDistDate_NRD,
    RegDistDate,
    RegDistDate_M,
    RegDistDate_NRD,
    RpState_NRD,
    RpRegOrg_NRD,
    PlacePrice_NRD,
    SumIssueVal,
    SumIssueVol,
    SumIssueVol_M,
    SumIssueVol_NRD,
    SumMarketVal,
    SumMarketVol,
    SumMarketVol_M,
    SumMarketVol_NRD,
    FirstCouponDate_NRD,
    BegMtyDate,
    PlanMtyDate_NRD,
    EndMtyDate,
    EndMtyDate_M,
    EndMtyDate_NRD,
    DaysAll,
    DaysAll_M,
    DaysAll_NRD)
VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
)""")

for r in range(1, sheet.nrows): #read loop for 2nd row
    ISIN_RegCode_NRDCode = sheet.cell(r,0).value
    FinToolType = sheet.cell(r,1).value
    Sec_Type_ID = sheet.cell(r,2).value
    SecTypeNameRus_NRD = sheet.cell(r,3).value
    SecTypeNameEng_NRD = sheet.cell(r,4).value
    Sec_Type_BR_Code = sheet.cell(r,5).value
    SecTypeNameBR_NRD = sheet.cell(r,6).value
    CFI = sheet.cell(r,7).value
    SecurityType = sheet.cell(r,8).value
    SecurityKind = sheet.cell(r,9).value
    Sec_Form_ID = sheet.cell(r,10).value
    SecFormNameRus_NRD = sheet.cell(r,11).value
    SecFormNameEng_NRD = sheet.cell(r,12).value
    CouponType = sheet.cell(r,13).value
    Rate_Type_ID = sheet.cell(r,14).value
    RateTypeNameRus_NRD = sheet.cell(r,15).value
    RateTypeNameEng_NRD = sheet.cell(r,16).value
    CouponTypeName_NRD = sheet.cell(r,17).value
    HaveOffer = sheet.cell(r,18).value
    AmortisedMty = sheet.cell(r,19).value
    MaturityGroup = sheet.cell(r,20).value
    IsConvertible = sheet.cell(r,21).value
    NickName = sheet.cell(r,22).value
    FullName = sheet.cell(r,23).value
    FullName_NRD = sheet.cell(r,24).value
    ISINCode = sheet.cell(r,25).value
    ISIN144A = sheet.cell(r,26).value
    NRDCode = sheet.cell(r,27).value
    RegCode = sheet.cell(r,28).value
    RegCode_M = sheet.cell(r,29).value
    RegCode_NRD = sheet.cell(r,30).value
    IssueNumber = sheet.cell(r,31).value
    BondSeries = sheet.cell(r,32).value
    FinToolID = sheet.cell(r,33).value
    ProgFinToolID = sheet.cell(r,34).value
    Status = sheet.cell(r,35).value
    Sec_State_ID = sheet.cell(r,36).value
    SecStateRus_NRD = sheet.cell(r,37).value
    SecStateEng_NRD = sheet.cell(r,38).value
    HaveDefault = sheet.cell(r,39).value
    GuaranteeType = sheet.cell(r,40).value
    GuaranteeAmount = sheet.cell(r,41).value
    BorrowerName = sheet.cell(r,42).value
    BorrowerOKPO = sheet.cell(r,43).value
    BorrowerSector = sheet.cell(r,44).value
    BorrowerUID = sheet.cell(r,45).value
    IssuerName = sheet.cell(r,46).value
    IssuerName_NRD = sheet.cell(r,47).value
    IssuerOKPO = sheet.cell(r,48).value
    IssuerUID = sheet.cell(r,49).value
    NumGuarantors = sheet.cell(r,50).value
    Acc_Open_Date_NRD = sheet.cell(r,51).value
    FacialAcc_NRD = sheet.cell(r,52).value
    RegistrarAccTypeDate_NRD = sheet.cell(r,53).value
    RegistrarAccType_NRD = sheet.cell(r,54).value
    Registrar_NRD = sheet.cell(r,55).value
    PrivateDist = sheet.cell(r,56).value
    Placement_Type_ID = sheet.cell(r,57).value
    PlacementType_NRD = sheet.cell(r,58).value
    RegDate = sheet.cell(r,59).value
    RegDate_M = sheet.cell(r,60).value
    RegDate_NRD = sheet.cell(r,61).value
    RegOrg = sheet.cell(r,62).value
    RegOrg_NRD = sheet.cell(r,63).value
    BegDistDate = sheet.cell(r,64).value
    BegDistDate_M = sheet.cell(r,65).value
    BegDistDate_NRD = sheet.cell(r,66).value
    EndDistDate = sheet.cell(r,67).value
    EndDistDate_M = sheet.cell(r,68).value
    EndDistDate_NRD = sheet.cell(r,69).value
    RegDistDate = sheet.cell(r,70).value
    RegDistDate_M = sheet.cell(r,71).value
    RegDistDate_NRD = sheet.cell(r,72).value
    RpState_NRD = sheet.cell(r,73).value
    RpRegOrg_NRD = sheet.cell(r,74).value
    PlacePrice_NRD = sheet.cell(r,75).value
    SumIssueVal = sheet.cell(r,76).value
    SumIssueVol = sheet.cell(r,77).value
    SumIssueVol_M = sheet.cell(r,78).value
    SumIssueVol_NRD = sheet.cell(r,79).value
    SumMarketVal = sheet.cell(r,80).value
    SumMarketVol = sheet.cell(r,81).value
    SumMarketVol_M = sheet.cell(r,82).value
    SumMarketVol_NRD = sheet.cell(r,83).value
    FirstCouponDate_NRD = sheet.cell(r,84).value
    BegMtyDate = sheet.cell(r,85).value
    PlanMtyDate_NRD = sheet.cell(r,86).value
    EndMtyDate = sheet.cell(r,87).value
    EndMtyDate_M = sheet.cell(r,88).value
    EndMtyDate_NRD = sheet.cell(r,89).value
    DaysAll = sheet.cell(r,90).value
    DaysAll_M = sheet.cell(r,91).value
    DaysAll_NRD = sheet.cell(r,92).value
    
    # Assign values from each row
    values = (ISIN_RegCode_NRDCode,
              FinToolType,
              Sec_Type_ID,
              SecTypeNameRus_NRD,
              SecTypeNameEng_NRD,
              Sec_Type_BR_Code,
              SecTypeNameBR_NRD,
              CFI,
              SecurityType,
              SecurityKind,
              Sec_Form_ID,
              SecFormNameRus_NRD,
              SecFormNameEng_NRD,
              CouponType,
              Rate_Type_ID,
              RateTypeNameRus_NRD,
              RateTypeNameEng_NRD,
              CouponTypeName_NRD,
              HaveOffer,
              AmortisedMty,
              MaturityGroup,
              IsConvertible,
              NickName,
              FullName,
              FullName_NRD,
              ISINCode,
              ISIN144A,
              NRDCode,
              RegCode,
              RegCode_M,
              RegCode_NRD,
              IssueNumber,
              BondSeries,
              FinToolID,
              ProgFinToolID,
              Status,
              Sec_State_ID,
              SecStateRus_NRD,
              SecStateEng_NRD,
              HaveDefault,
              GuaranteeType,
              GuaranteeAmount,
              BorrowerName,
              BorrowerOKPO,
              BorrowerSector,
              BorrowerUID,
              IssuerName,
              IssuerName_NRD,
              IssuerOKPO,
              IssuerUID,
              NumGuarantors,
              Acc_Open_Date_NRD,
              FacialAcc_NRD,
              RegistrarAccTypeDate_NRD,
              RegistrarAccType_NRD,
              Registrar_NRD,
              PrivateDist,
              Placement_Type_ID,
              PlacementType_NRD,
              RegDate,
              RegDate_M,
              RegDate_NRD,
              RegOrg,
              RegOrg_NRD, 
              BegDistDate,
              BegDistDate_M,
              BegDistDate_NRD,
              EndDistDate,
              EndDistDate_M,
              EndDistDate_NRD,
              RegDistDate,
              RegDistDate_M,
              RegDistDate_NRD,
              RpState_NRD,
              RpRegOrg_NRD,
              PlacePrice_NRD,
              SumIssueVal,
              SumIssueVol,
              SumIssueVol_M,
              SumIssueVol_NRD,
              SumMarketVal,
              SumMarketVol,
              SumMarketVol_M,
              SumMarketVol_NRD,
              FirstCouponDate_NRD,
              BegMtyDate,
              PlanMtyDate_NRD,
              EndMtyDate,
              EndMtyDate_M,
              EndMtyDate_NRD,
              DaysAll,
              DaysAll_M,
              DaysAll_NRD)
    cursor.execute(query2, values)
connection.commit()

cursor.execute("""
CREATE TABLE [dbo].[Base_1](
    ID varchar(80),
    TIME varchar(80), 
    ACCRUEDINT varchar(80),
    ASK varchar(80),
    ASK_SIZE varchar(80),
    ASK_SIZE_TOTAL varchar(80), 
    AVGE_PRCE varchar(80),
    BID varchar(80),
    BID_SIZE varchar(80), 
    BID_SIZE_TOTAL varchar(80),
    BOARDID varchar(80),
    BOARDNAME varchar(80),
    BUYBACKDATE varchar(80),
    BUYBACKPRICE varchar(80),
    CBR_LOMBARD varchar(80),
    CBR_PLEDGE varchar(80),
    CLOSE_ varchar(80),
    CPN varchar(80),
    CPN_DATE varchar(80),
    CPN_PERIOD varchar(80),
    DEAL_ACC varchar(80),
    FACEVALUE varchar(80),
    ISIN varchar(80),
    ISSUER varchar(80),
    ISSUESIZE varchar(80), 
    MAT_DATE varchar(80),
    MPRICE varchar(80),
    MPRICE2 varchar(80),
    SPREAD varchar(80),
    VOL_ACC varchar(80),
    Y2O_ASK varchar(80),
    Y2O_BID varchar(80),
    YIELD_ASK varchar(80),
    YIELD_BID varchar(80)
)""")
connection.commit()

data = pd.read_excel('bp.xlsx') # BASE_PRICES (BASE1)
data = data.rename(columns={'ID': 'ID',
                            'TIME': 'TIME', 
                            'ACCRUEDINT': 'ACCRUEDINT',
                            'ASK': 'ASK',
                            'ASK_SIZE': 'ASK_ASK_SIZE',
                            'ASK_SIZE_TOTAL': 'ASK_SIZE_TOTAL', 
                            'AVGE_PRCE': 'AVGE_PRCE',
                            'BID': 'BID',                            
                            'BID_SIZE': 'BID_SIZE', 
                            'BID_SIZE_TOTAL': 'BID_SIZE_TOTAL',
                            'BOARDID': 'BOARDID',
                            'BOARDNAME': 'BOARDNAME',
                            'BUYBACKDATE': 'BUYBACKDATE',
                            'BUYBACKPRICE': 'BUYBACKPRICE',
                            'CBR_LOMBARD': 'CBR_LOMBARD',
                            'CBR_PLEDGE': 'CBR_PLEDGE',
                            'CLOSE': 'CLOSE_',
                            'CPN': 'CPN',
                            'CPN_DATE': 'CPN_DATE',
                            'CPN_PERIOD': 'CPN_PERIOD',
                            'DEAL_ACC': 'DEAL_ACC',
                            'FACEVALUE': 'FACEVALUE',
                            'ISIN': 'ISIN',
                            'ISSUER': 'ISSUER',
                            'ISSUESIZE': 'ISSUESIZE',
                            'MAT_DATE': 'MAT_DATE',
                            'MPRICE': 'MPRICE',
                            'MPRICE2': 'MPRICE2',
                            'SPREAD': 'SPREAD',
                            'VOL_ACC': 'VOL_ACC',
                            'Y2O_ASK': 'Y2O_ASK',
                            'Y2O_BID': 'Y2O_BID',                                                  
                            'YIELD_ASK': 'Y2O_ASK',
                            'YIELD_BID': 'Y2O_ASK'}) # rename columns

# export
data.to_excel('bp.xlsx', index=False)

# Open the workbook and define the worksheet
book = xlrd.open_workbook('bp.xlsx')
sheet = book.sheet_by_name('Sheet1')

query3 = ("""
INSERT INTO [Risks].[dbo].[Base_1](
    ID,
    TIME, 
    ACCRUEDINT,
    ASK,
    ASK_SIZE,
    ASK_SIZE_TOTAL, 
    AVGE_PRCE,
    BID,
    BID_SIZE, 
    BID_SIZE_TOTAL,
    BOARDID,
    BOARDNAME,
    BUYBACKDATE,
    BUYBACKPRICE,
    CBR_LOMBARD,
    CBR_PLEDGE,
    CLOSE_,
    CPN,
    CPN_DATE,
    CPN_PERIOD,
    DEAL_ACC,
    FACEVALUE,
    ISIN,
    ISSUER,
    ISSUESIZE,
    MAT_DATE,
    MPRICE,
    MPRICE2,
    SPREAD,
    VOL_ACC,
    Y2O_ASK,
    Y2O_BID,
    YIELD_ASK,
    YIELD_BID) 
VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
)""")

for r in range(1, sheet.nrows): #read loop for 2nd row
    ID = sheet.cell(r,0).value
    TIME = sheet.cell(r,1).value
    ACCRUEDINT = sheet.cell(r,2).value
    ASK = sheet.cell(r,3).value
    ASK_SIZE = sheet.cell(r,4).value
    ASK_SIZE_TOTAL = sheet.cell(r,5).value
    AVGE_PRCE = sheet.cell(r,6).value
    BID = sheet.cell(r,7).value
    BID_SIZE = sheet.cell(r,8).value
    BID_SIZE_TOTAL = sheet.cell(r,9).value
    BOARDID = sheet.cell(r,10).value
    BOARDNAME = sheet.cell(r,11).value
    BUYBACKDATE = sheet.cell(r,12).value
    BUYBACKPRICE = sheet.cell(r,13).value
    CBR_LOMBARD = sheet.cell(r,14).value
    CBR_PLEDGE = sheet.cell(r,15).value
    CLOSE_ = sheet.cell(r,16).value
    CPN = sheet.cell(r,17).value
    CPN_DATE = sheet.cell(r,18).value
    CPN_PERIOD = sheet.cell(r,19).value
    DEAL_ACC = sheet.cell(r,20).value
    FACEVALUE = sheet.cell(r,21).value
    ISIN = sheet.cell(r,22).value
    ISSUER = sheet.cell(r,23).value
    ISSUESIZE = sheet.cell(r,24).value
    MAT_DATE = sheet.cell(r,25).value
    MPRICE = sheet.cell(r,26).value
    MPRICE2 = sheet.cell(r,27).value
    SPREAD = sheet.cell(r,28).value
    VOL_ACC = sheet.cell(r,29).value
    Y2O_ASK = sheet.cell(r,30).value
    Y2O_BID = sheet.cell(r,31).value
    YIELD_ASK = sheet.cell(r,32).value
    YIELD_BID = sheet.cell(r,33).value
    
    # Assign values from each row
    values = (ID,
              TIME,
              ACCRUEDINT,
              ASK,
              ASK_SIZE,
              ASK_SIZE_TOTAL, 
              AVGE_PRCE,
              BID,
              BID_SIZE, 
              BID_SIZE_TOTAL,
              BOARDID,
              BOARDNAME,
              BUYBACKDATE,
              BUYBACKPRICE,
              CBR_LOMBARD,
              CBR_PLEDGE,
              CLOSE_,
              CPN,
              CPN_DATE,
              CPN_PERIOD,
              DEAL_ACC,
              FACEVALUE,
              ISIN,
              ISSUER,
              ISSUESIZE,
              MAT_DATE,
              MPRICE,
              MPRICE2,
              SPREAD,
              VOL_ACC,
              Y2O_ASK,
              Y2O_BID,
              YIELD_ASK,
              YIELD_BID)
    cursor.execute(query3, values)
connection.commit()

cursor.execute("""
CREATE TABLE [dbo].[Base_2](
    ID varchar(80),
    TIME varchar(80), 
    ACCRUEDINT varchar(80),
    ASK varchar(80),
    ASK_SIZE varchar(80),
    ASK_SIZE_TOTAL varchar(80), 
    AVGE_PRCE varchar(80),
    BID varchar(80),
    BID_SIZE varchar(80), 
    BID_SIZE_TOTAL varchar(80),
    BOARDID varchar(80),
    BOARDNAME varchar(80),
    BUYBACKDATE varchar(80),
    BUYBACKPRICE varchar(80),
    CBR_LOMBARD varchar(80),
    CBR_PLEDGE varchar(80),
    CLOSE_ varchar(80),
    CPN varchar(80),
    CPN_DATE varchar(80),
    CPN_PERIOD varchar(80),
    DEAL_ACC varchar(80),
    FACEVALUE varchar(80),
    ISIN varchar(80),
    ISSUER varchar(80),
    ISSUESIZE varchar(80), 
    MAT_DATE varchar(80),
    MPRICE varchar(80),
    MPRICE2 varchar(80),
    SPREAD varchar(80),
    VOL_ACC varchar(80),
    Y2O_ASK varchar(80),
    Y2O_BID varchar(80),
    YIELD_ASK varchar(80),
    YIELD_BID varchar(80)
)""")
connection.commit()

data = pd.read_excel('bp2.xlsx') # BASE_PRICES (BASE2)
data = data.rename(columns={'ID': 'ID',
                            'TIME': 'TIME', 
                            'ACCRUEDINT': 'ACCRUEDINT',
                            'ASK': 'ASK',
                            'ASK_SIZE': 'ASK_ASK_SIZE',
                            'ASK_SIZE_TOTAL': 'ASK_SIZE_TOTAL', 
                            'AVGE_PRCE': 'AVGE_PRCE',
                            'BID': 'BID',                            
                            'BID_SIZE': 'BID_SIZE', 
                            'BID_SIZE_TOTAL': 'BID_SIZE_TOTAL',
                            'BOARDID': 'BOARDID',
                            'BOARDNAME': 'BOARDNAME',
                            'BUYBACKDATE': 'BUYBACKDATE',
                            'BUYBACKPRICE': 'BUYBACKPRICE',
                            'CBR_LOMBARD': 'CBR_LOMBARD',
                            'CBR_PLEDGE': 'CBR_PLEDGE',
                            'CLOSE': 'CLOSE_',
                            'CPN': 'CPN',
                            'CPN_DATE': 'CPN_DATE',
                            'CPN_PERIOD': 'CPN_PERIOD',
                            'DEAL_ACC': 'DEAL_ACC',
                            'FACEVALUE': 'FACEVALUE',
                            'ISIN': 'ISIN',
                            'ISSUER': 'ISSUER',
                            'ISSUESIZE': 'ISSUESIZE',
                            'MAT_DATE': 'MAT_DATE',
                            'MPRICE': 'MPRICE',
                            'MPRICE2': 'MPRICE2',
                            'SPREAD': 'SPREAD',
                            'VOL_ACC': 'VOL_ACC',
                            'Y2O_ASK': 'Y2O_ASK',
                            'Y2O_BID': 'Y2O_BID',                                                  
                            'YIELD_ASK': 'Y2O_ASK',
                            'YIELD_BID': 'Y2O_ASK'}) # rename columns

# export
data.to_excel('bp2.xlsx', index=False)

# Open the workbook and define the worksheet
book = xlrd.open_workbook('bp2.xlsx')
sheet = book.sheet_by_name('Sheet1')

query4 = ("""
INSERT INTO [Risks].[dbo].[Base_2](
    ID,
    TIME, 
    ACCRUEDINT,
    ASK,
    ASK_SIZE,
    ASK_SIZE_TOTAL, 
    AVGE_PRCE,
    BID,
    BID_SIZE, 
    BID_SIZE_TOTAL,
    BOARDID,
    BOARDNAME,
    BUYBACKDATE,
    BUYBACKPRICE,
    CBR_LOMBARD,
    CBR_PLEDGE,
    CLOSE_,
    CPN,
    CPN_DATE,
    CPN_PERIOD,
    DEAL_ACC,
    FACEVALUE,
    ISIN,
    ISSUER,
    ISSUESIZE,
    MAT_DATE,
    MPRICE,
    MPRICE2,
    SPREAD,
    VOL_ACC,
    Y2O_ASK,
    Y2O_BID,
    YIELD_ASK,
    YIELD_BID) 
VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
)""")

for r in range(1, sheet.nrows): #read loop for 2nd row
    ID = sheet.cell(r,0).value
    TIME = sheet.cell(r,1).value
    ACCRUEDINT = sheet.cell(r,2).value
    ASK = sheet.cell(r,3).value
    ASK_SIZE = sheet.cell(r,4).value
    ASK_SIZE_TOTAL = sheet.cell(r,5).value
    AVGE_PRCE = sheet.cell(r,6).value
    BID = sheet.cell(r,7).value
    BID_SIZE = sheet.cell(r,8).value
    BID_SIZE_TOTAL = sheet.cell(r,9).value
    BOARDID = sheet.cell(r,10).value
    BOARDNAME = sheet.cell(r,11).value
    BUYBACKDATE = sheet.cell(r,12).value
    BUYBACKPRICE = sheet.cell(r,13).value
    CBR_LOMBARD = sheet.cell(r,14).value
    CBR_PLEDGE = sheet.cell(r,15).value
    CLOSE_ = sheet.cell(r,16).value
    CPN = sheet.cell(r,17).value
    CPN_DATE = sheet.cell(r,18).value
    CPN_PERIOD = sheet.cell(r,19).value
    DEAL_ACC = sheet.cell(r,20).value
    FACEVALUE = sheet.cell(r,21).value
    ISIN = sheet.cell(r,22).value
    ISSUER = sheet.cell(r,23).value
    ISSUESIZE = sheet.cell(r,24).value
    MAT_DATE = sheet.cell(r,25).value
    MPRICE = sheet.cell(r,26).value
    MPRICE2 = sheet.cell(r,27).value
    SPREAD = sheet.cell(r,28).value
    VOL_ACC = sheet.cell(r,29).value
    Y2O_ASK = sheet.cell(r,30).value
    Y2O_BID = sheet.cell(r,31).value
    YIELD_ASK = sheet.cell(r,32).value
    YIELD_BID = sheet.cell(r,33).value
    
    # Assign values from each row
    values = (ID,
              TIME,
              ACCRUEDINT,
              ASK,
              ASK_SIZE,
              ASK_SIZE_TOTAL, 
              AVGE_PRCE,
              BID,
              BID_SIZE, 
              BID_SIZE_TOTAL,
              BOARDID,
              BOARDNAME,
              BUYBACKDATE,
              BUYBACKPRICE,
              CBR_LOMBARD,
              CBR_PLEDGE,
              CLOSE_,
              CPN,
              CPN_DATE,
              CPN_PERIOD,
              DEAL_ACC,
              FACEVALUE,
              ISIN,
              ISSUER,
              ISSUESIZE,
              MAT_DATE,
              MPRICE,
              MPRICE2,
              SPREAD,
              VOL_ACC,
              Y2O_ASK,
              Y2O_BID,
              YIELD_ASK,
              YIELD_BID)
    cursor.execute(query4, values)
connection.commit()

cursor.execute("""
CREATE TABLE [dbo].[RiskFree](
    DATE varchar(10),
    RFR1Q varchar(10), 
    RFRHY varchar(10),
    RFR3Q varchar(10),
    RFR1 varchar(10),
    RFR2 varchar(10), 
    RFR3 varchar(10),
    RFR5 varchar(10),
    RFR7 varchar(10), 
    RFR10 varchar(10),
    RFR15 varchar(10),
    RFR20 varchar(10),
    RFR30 varchar(10))""")
connection.commit()

data = pd.read_excel('rfr.xlsx') # RISKFREERATES
data = data.rename(columns={'DATE': 'DATE',
                            'RFR1Q': 'RFR1Q',
                            'RFRHY': 'RFRHY',
                            'RFR3Q': 'RFR3Q',
                            'RFR1': 'RFR1',
                            'RFR2': 'RFR2',
                            'RFR3': 'RFR3',
                            'RFR5': 'RFR5',
                            'RFR7': 'RFR7',
                            'RFR10': 'RFR10',
                            'RFR15': 'RFR15',
                            'RFR20': 'RFR20',
                            'RFR30': 'RFR30'}) # rename columns

# export
data.to_excel('rfr.xlsx', index=False)

# Open the workbook and define the worksheet
book = xlrd.open_workbook('rfr.xlsx')
sheet = book.sheet_by_name('Sheet1')

query5 = ("""
INSERT INTO [Risks].[dbo].[RiskFree](
    DATE,
    RFR1Q,
    RFRHY,
    RFR3Q,
    RFR1,
    RFR2,
    RFR3,
    RFR5,
    RFR7,
    RFR10,
    RFR15,
    RFR20,
    RFR30) 
VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?
)""")

for r in range(1, sheet.nrows): #read loop for 2nd row
    DATE = sheet.cell(r,0).value
    RFR1Q = sheet.cell(r,1).value
    RFRHY = sheet.cell(r,2).value
    RFR3Q = sheet.cell(r,3).value
    RFR1 = sheet.cell(r,4).value
    RFR2 = sheet.cell(r,5).value
    RFR3 = sheet.cell(r,6).value
    RFR5 = sheet.cell(r,7).value
    RFR7 = sheet.cell(r,8).value
    RFR10 = sheet.cell(r,9).value
    RFR15 = sheet.cell(r,10).value
    RFR20 = sheet.cell(r,11).value
    RFR30 = sheet.cell(r,12).value
    
    # Assign values from each row
    values = (DATE,
              RFR1Q,
              RFRHY,
              RFR3Q,
              RFR1,
              RFR2,
              RFR3,
              RFR5,
              RFR7,
              RFR10,
              RFR15,
              RFR20,
              RFR30)
    cursor.execute(query5, values)
connection.commit()
# Close special object
cursor.close()
# Close the database connection
connection.close()

