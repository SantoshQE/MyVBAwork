import xlrd
import pandas as pd
import os


#*******************************************************************************************************
 #   'Read the input excel sheet
#*******************************************************************************************************
os.chdir("..")
inputFilePath =  str(os.path.abspath(os.curdir)) + "\SourceVsOutboundValidator_V1.0.xlsx"
print('InputFilePath directory is : ' + inputFilePath)
#*******************************************************************************************************
# Give the location of the file
#*******************************************************************************************************
loc = (inputFilePath)
inputWB = xlrd.open_workbook(loc)
inputSheet = inputWB.sheet_by_name("Inputs")
SMD_File_Name= str(os.path.abspath(os.curdir)) + str("\\") + str((inputSheet.cell_value(2, 1)))
Business_Rule_Column_Name =  str((inputSheet.cell_value(3, 1)))
ProductName = (inputSheet.cell_value(4, 1))
Sub_Product_Name = (inputSheet.cell_value(5, 1))
Source_File_Name = (inputSheet.cell_value(6, 1))
Product_Col_Name = (inputSheet.cell_value(7, 1))
print('Start Reading Read SMD...')
smdWB = xlrd.open_workbook(SMD_File_Name)
smdSheet = smdWB.sheet_by_name("Customized")
i = 0
#************* FUNCTION TO GET THE COLUMN INDEX USING COLUMN NAME **************************************
def ReturnExcelColIndex(colname):
    for i in range(smdSheet.ncols):
        if smdSheet.cell_value(1, i) == colname:
            #continue
            ColIndex = i
            print("ColIndex..." + str(i))
            return ColIndex;
            break
#*******************************************************************************************************
 #   'Get the SMD column headers to continue further process
#*******************************************************************************************************
SMD_ProductNameColIndex = ReturnExcelColIndex(colname = "ProductName")
SMD_SubProductNameColIndex = ReturnExcelColIndex(colname = "SubProductName")
SMD_FileTypeColIndex = ReturnExcelColIndex(colname = "FileType")
SMD_FileAttributeColIndex = ReturnExcelColIndex(colname = "FileAttribute")
SMD_DataTypeColIndex = ReturnExcelColIndex(colname = "DataType")
SMD_OtisEnumTypeColIndex = ReturnExcelColIndex(colname = "OtisEnumType")
SMD_MandatoryStatusColIndex = ReturnExcelColIndex(colname = "MandatoryStatus")
SMD_EODcBusinessKeyColIndex = ReturnExcelColIndex(colname = "EODcBusinessKey")
SMD_BusinessRuleColIndex = ReturnExcelColIndex(colname = Product_Col_Name +"." +"Business Rule")
#*******************************************************************************************************
#*******************************************************************************************************

#*******************************************************************************************************
# FUNCTION TO LOAD THE OUTBOUND PSV FILES
#*******************************************************************************************************
def LoadOutBoundPSV(eodFileType):
    n=11
    for n in range(smdSheet.ncols):
        if inputSheet.cell_value(n, 2) == eodFileType:
            data = pd.read_csv(str(os.path.abspath(os.curdir)) + "\\" + str(inputSheet.cell_value(n, 1)), sep='|',skiprows=[0])
            df = pd.DataFrame(data)
            header_list = list(df.columns)
            #print(header_list)
            return header_list;

m=5
#in range(smdSheet.nrows):
for m in range(smdSheet.nrows):
    SMD_productVal = smdSheet.cell_value(m, SMD_ProductNameColIndex)
    SMD_subproductVal = smdSheet.cell_value(m, SMD_SubProductNameColIndex)
    SMD_FileTypeVal = smdSheet.cell_value(m, SMD_FileTypeColIndex)
    SMD_FileAttributeVal = smdSheet.cell_value(m, SMD_FileAttributeColIndex)
    SMD_DataTypeVal = smdSheet.cell_value(m, SMD_DataTypeColIndex)
    SMD_OtisEnumTypeVal = smdSheet.cell_value(m, SMD_OtisEnumTypeColIndex)
    SMD_MandatoryStatusVal = smdSheet.cell_value(m, SMD_MandatoryStatusColIndex)
    SMD_EODcBusinessKeyVal = smdSheet.cell_value(m, SMD_EODcBusinessKeyColIndex)
    SMD_BusinessRuleVal = smdSheet.cell_value(m, SMD_BusinessRuleColIndex)

    if(SMD_BusinessRuleVal == "") and (SMD_DataTypeVal != "ENUM"):
            print(SMD_FileAttributeVal + "skipped from data validation")
    elif (SMD_BusinessRuleVal != "") and SMD_BusinessRuleVal == "Direct Move":
        # READ THE PRIMARY KEY FROM BOTH THE OUTBOUND AND SOURCE FILE
        if(SMD_FileTypeVal =="EODcPosition"):
            LoadOutBoundPSV("EODcPosition")  # LOADED THE OUTBOUND FILE AND HEADER LIST
            # iterating over rows using iterrows() function
           # print(data.head(3))




