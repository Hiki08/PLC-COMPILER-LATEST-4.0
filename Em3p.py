#%%
from Imports import *
import DateAndTimeManager
from FilesReader import *

#%%
class em3P():
    em3PData = ""
    em3PItemCode = ""

    totalAverage3 = []
    totalAverage4 = []
    totalAverage5 = []
    totalAverage10 = []

    totalMinimum3 = []
    totalMinimum4 = []
    totalMinimum5 = []
    
    totalMaximum3 = []
    totalMaximum4 = []
    totalMaximum5 = []

    readingYear = ""
    fileFinishedReading = False
    fileList = []

    isValueRetrieve = False

    def __init__(self):
        pass

    def GettingData(self, itemCode, lotNumber):
        if itemCode == "EM0580107P":
            self.fileList = EM0580107PData
        elif itemCode == "EM0660047P":
            self.fileList = EM0660047PData
        elif itemCode == "EM0660045P":
            self.fileList = EM0660045PData

        for fileNum in range(len(self.fileList)):
            self.totalAverage3 = []
            self.totalAverage4 = []
            self.totalAverage5 = []
            self.totalAverage10 = []

            self.totalMinimum3 = []
            self.totalMinimum4 = []
            self.totalMinimum5 = []
            
            self.totalMaximum3 = []
            self.totalMaximum4 = []
            self.totalMaximum5 = []
            
            print(f"READING FILE {fileNum}")

            #Getting The Row, Column Location Of HIBLOW
            findSupplier = [(index, column) for index, row in self.fileList[fileNum].iterrows() for column, value in row.items() if value == "SUPPLIER"]
            supplierRow = [index for index, _ in findSupplier]
            supplierColumn = [column for _, column in findSupplier]

            print("Row indices:", supplierRow)
            print("Column names:", supplierColumn)

            # Get the Neighboring Data Of SUPPLIER
            supplierFiltered = self.fileList[fileNum].iloc[max(0, supplierRow[0] - 3):min(len(self.fileList[fileNum]), supplierRow[0] + 10), self.fileList[fileNum].columns.get_loc(supplierColumn[0]):self.fileList[fileNum].columns.get_loc(supplierColumn[0]) + 999999]
            supplierFiltered

            try:
                #Getting The Row, Column Location Of Lot Number
                findLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == lotNumber]
                lotNumberRow = [index for index, _ in findLotNumber]
                lotNumberColumn = [column for _, column in findLotNumber]

                print("Row indices:", lotNumberRow)
                print("Column names:", lotNumberColumn)

                for a in range(0, len(lotNumberColumn)):
                    # Get The Neighboring Data of Lot Number
                    inspectionData = self.fileList[fileNum].iloc[max(0, lotNumberRow[a]):min(len(self.fileList[fileNum]), lotNumberRow[a] + 13), self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]):self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]) + 5]

                    #CHECKING THE ITEM CODE
                    if itemCode == "EM0580107P":
                        average3 = inspectionData.iloc[5].mean()
                        average4 = inspectionData.iloc[6].mean()
                        average5 = inspectionData.iloc[7].mean()
                        average10 = inspectionData.iloc[12, 0]

                        print(f"Inspection 10: {average10}")
                        # if not average10 is np.nan:
                            

                        minimum3 = inspectionData.iloc[5].min()
                        minimum4 = inspectionData.iloc[6].min()
                        minimum5 = inspectionData.iloc[7].min()

                        maximum3 = inspectionData.iloc[5].max()
                        maximum4 = inspectionData.iloc[6].max()
                        maximum5 = inspectionData.iloc[7].max()

                        self.totalAverage3.append(average3)
                        self.totalAverage4.append(average4)
                        self.totalAverage5.append(average5)
                        self.totalAverage10.append(average10)

                        self.totalMinimum3.append(minimum3)
                        self.totalMinimum4.append(minimum4)
                        self.totalMinimum5.append(minimum5)

                        self.totalMaximum3.append(maximum3)
                        self.totalMaximum4.append(maximum4)
                        self.totalMaximum5.append(maximum5)
                    #CHECKING THE ITEM CODE
                    elif itemCode == "EM0660047P":
                        average3 = inspectionData.iloc[5].mean()
                        average10 = inspectionData.iloc[8, 0]

                        minimum3 = inspectionData.iloc[5].min()
                    
                        maximum3 = inspectionData.iloc[5].max()
                        
                        self.totalAverage3.append(average3)
                        self.totalAverage10.append(average10)

                        self.totalMinimum3.append(minimum3)

                        self.totalMaximum3.append(maximum3)
                    #CHECKING THE ITEM CODE
                    elif itemCode == "EM0660045P":
                        average3 = inspectionData.iloc[5].mean()
                        average10 = inspectionData.iloc[8, 0]

                        minimum3 = inspectionData.iloc[5].min()
                    
                        maximum3 = inspectionData.iloc[5].max()
                        
                        self.totalAverage3.append(average3)
                        self.totalAverage10.append(average10)

                        self.totalMinimum3.append(minimum3)

                        self.totalMaximum3.append(maximum3)
                    
                #CHECKING THE ITEM CODE
                if itemCode == "EM0580107P":
                    self.totalAverage3 = statistics.mean(self.totalAverage3)
                    self.totalAverage4 = statistics.mean(self.totalAverage4)
                    self.totalAverage5 = statistics.mean(self.totalAverage5)
                    self.totalAverage10 = statistics.mean(self.totalAverage10)

                    self.totalMinimum3 = min(self.totalMinimum3)
                    self.totalMinimum4 = min(self.totalMinimum4)
                    self.totalMinimum5 = min(self.totalMinimum5)

                    self.totalMaximum3 = max(self.totalMaximum3)
                    self.totalMaximum4 = max(self.totalMaximum4)
                    self.totalMaximum5 = max(self.totalMaximum5)

                    self.totalAverage3 = f"{self.totalAverage3:.2f}"
                    self.totalAverage4 = f"{self.totalAverage4:.2f}"
                    self.totalAverage5 = f"{self.totalAverage5:.2f}"
                    self.totalAverage10 = f"{self.totalAverage10:.2f}"

                    self.totalMinimum3 = f"{self.totalMinimum3:.2f}"
                    self.totalMinimum4 = f"{self.totalMinimum4:.2f}"
                    self.totalMinimum5 = f"{self.totalMinimum5:.2f}"

                    self.totalMaximum3 = f"{self.totalMaximum3:.2f}"
                    self.totalMaximum4 = f"{self.totalMaximum4:.2f}"
                    self.totalMaximum5 = f"{self.totalMaximum5:.2f}"

                    break
                #CHECKING THE ITEM CODE
                elif itemCode == "EM0660047P":
                    self.totalAverage3 = statistics.mean(self.totalAverage3)
                    self.totalAverage10 = statistics.mean(self.totalAverage10)

                    self.totalMinimum3 = min(self.totalMinimum3)

                    self.totalMaximum3 = max(self.totalMaximum3)

                    self.totalAverage3 = f"{self.totalAverage3:.2f}"
                    self.totalAverage10 = f"{self.totalAverage10:.2f}"

                    self.totalMinimum3 = f"{self.totalMinimum3:.2f}"

                    self.totalMaximum3 = f"{self.totalMaximum3:.2f}"
                    break
                #CHECKING THE ITEM CODE
                elif itemCode == "EM0660045P":
                    self.totalAverage3 = statistics.mean(self.totalAverage3)
                    self.totalAverage10 = statistics.mean(self.totalAverage10)

                    self.totalMinimum3 = min(self.totalMinimum3)

                    self.totalMaximum3 = max(self.totalMaximum3)

                    self.totalAverage3 = f"{self.totalAverage3:.2f}"
                    self.totalAverage10 = f"{self.totalAverage10:.2f}"

                    self.totalMinimum3 = f"{self.totalMinimum3:.2f}"

                    self.totalMaximum3 = f"{self.totalMaximum3:.2f}"
                    break
            except:
                self.totalAverage3 = "No Data Found"
                self.totalAverage4 = "No Data Found"
                self.totalAverage5 = "No Data Found"
                self.totalAverage10 = "No Data Found"

                self.totalMinimum3 = "No Data Found"
                self.totalMinimum4 = "No Data Found"
                self.totalMinimum5 = "No Data Found"

                self.totalMaximum3 = "No Data Found"
                self.totalMaximum4 = "No Data Found"
                self.totalMaximum5 = "No Data Found"

            print(f"Total Average: {self.totalAverage3}")
            print(f"Total Average: {self.totalAverage4}")
            print(f"Total Average: {self.totalAverage5}")
            print(f"Total Average: {self.totalAverage10}")

            print(f"Total Minimum: {self.totalMinimum3}")
            print(f"Total Minimum: {self.totalMinimum4}")
            print(f"Total Minimum: {self.totalMinimum5}")

            print(f"Total Maximum: {self.totalMaximum3}")
            print(f"Total Maximum: {self.totalMaximum4}")
            print(f"Total Maximum: {self.totalMaximum5}")

        print(f"Selected Total Average: {self.totalAverage3}")
        print(f"Selected Total Average: {self.totalAverage4}")
        print(f"Selected Total Average: {self.totalAverage5}")
        print(f"Selected Total Average: {self.totalAverage10}")

        print(f"Selected Total Minimum: {self.totalMinimum3}")
        print(f"Selected Total Minimum: {self.totalMinimum4}")
        print(f"Selected Total Minimum: {self.totalMinimum5}")

        print(f"Selected Total Maximum: {self.totalMaximum3}")
        print(f"Selected Total Maximum: {self.totalMaximum4}")
        print(f"Selected Total Maximum: {self.totalMaximum5}")

    def Trial(self, lotNumber):
        fileNum = 0
        a = 2

        #Getting The Row, Column Location Of HIBLOW
        findSupplier = [(index, column) for index, row in self.fileList[fileNum].iterrows() for column, value in row.items() if value == "SUPPLIER"]
        supplierRow = [index for index, _ in findSupplier]
        supplierColumn = [column for _, column in findSupplier]

        print("Row indices:", supplierRow)
        print("Column names:", supplierColumn)

        # Get the Neighboring Data Of SUPPLIER
        supplierFiltered = self.fileList[fileNum].iloc[max(0, supplierRow[0] - 3):min(len(self.fileList[fileNum]), supplierRow[0] + 10), self.fileList[fileNum].columns.get_loc(supplierColumn[0]):self.fileList[fileNum].columns.get_loc(supplierColumn[0]) + 999999]
        
        #Getting The Row, Column Location Of Lot Number
        findLotNumber = [(index, column) for index, row in supplierFiltered.iterrows() for column, value in row.items() if value == lotNumber]
        lotNumberRow = [index for index, _ in findLotNumber]
        lotNumberColumn = [column for _, column in findLotNumber]

        print("Row indices:", lotNumberRow)
        print("Column names:", lotNumberColumn)

        inspectionData = self.fileList[fileNum].iloc[max(0, lotNumberRow[a]):min(len(self.fileList[fileNum]), lotNumberRow[a] + 13), self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]):self.fileList[fileNum].columns.get_loc(lotNumberColumn[a]) + 5]
        return inspectionData
    
#%%
# em3p = em3P()
# DateAndTimeManager.GetDateToday()
# em3p.readingYear = int(DateAndTimeManager.yearNow)
# em3p.ReadExcel("EM0580107P")
# # em3p.ReadExcel("EM0660047P")
# # em3p.ReadExcel("EM0660045P")
# print(f"Total Number Of Files {len(em3p.fileList)}")
# em3p.GettingData("CAT-5A07DI")
# # em3p.GettingData("FC6030-4F05GT")
# # em3p.GettingData("FC6030-4G26GT")
# # em3p.Trial("CAT-5A07DI")
# %%
